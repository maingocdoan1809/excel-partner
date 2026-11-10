import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import store from "./runtime-store.cjs";
import { installPackage, installRuntime } from "./runtime-manager.mjs";
import { runCommand } from "./doctor-probes.mjs";

const fileHash = file => store.fileDigest(file);
function inventory(directory) {
  const files = {};
  function walk(current) {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const file = path.join(current, entry.name), relative = path.relative(directory, file);
      if (entry.name === ".release.json") continue;
      if (entry.isSymbolicLink()) {
        const real = fs.realpathSync(file);
        if (!store.inside(directory, real)) throw new Error("ARTIFACT_EXTERNAL_LINK");
        files[relative] = { link: path.relative(directory, real) }; continue;
      }
      if (entry.isDirectory()) walk(file);
      else files[relative] = { sha256: fileHash(file) };
    }
  }
  walk(directory); return files;
}
export function verifyRelease(release, deep = false) {
  if (!fs.existsSync(path.join(release.root, "package.json"))) throw new Error("RELEASE_MISSING");
  const record = store.readJson(path.join(release.directory, ".release.json"), 64 * 1024 * 1024);
  const pkg = store.readJson(path.join(release.root, "package.json"));
  if (pkg.name !== "office-partner" || pkg.version !== release.version || record.id !== release.id) throw new Error("RELEASE_IDENTITY_MISMATCH");
  const checks = deep ? Object.keys(record.files) : [path.relative(release.directory, path.join(release.root, "bin", "office-partner.mjs")),
    path.relative(release.directory, path.join(release.root, "dist", "bridge", "src", "index.js")),
    path.relative(release.directory, path.join(release.root, "gateway", "dist", "server.mjs"))];
  for (const relative of checks) {
    const expected = record.files[relative], file = path.join(release.directory, relative);
    if (!expected || !store.inside(release.directory, file)) throw new Error("RELEASE_INTEGRITY_FAILED");
    if (expected.link) {
      if (!store.inside(release.directory, fs.realpathSync(file)) || path.relative(release.directory, fs.realpathSync(file)) !== expected.link) throw new Error("RELEASE_INTEGRITY_FAILED");
    } else if (!fs.existsSync(file) || fileHash(file) !== expected.sha256) throw new Error("RELEASE_INTEGRITY_FAILED");
  }
  return true;
}
function remember(dataDir, release) {
  const deployment = store.currentDeployment(dataDir); deployment.releases[release.id] = release;
  store.atomic(path.join(dataDir, "deployment.json"), deployment);
}
function finalize(directory, root, dataDir, id) {
  const pkg = store.readJson(path.join(root, "package.json"));
  if (pkg.name !== "office-partner" || !/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(pkg.version)) throw new Error("RELEASE_PACKAGE_INVALID");
  for (const required of ["bin/office-partner.mjs", "bin/runtime-store.cjs", "bin/runtime-manifest.json", "dist/bridge/src/index.js", "gateway/dist/server.mjs", "addin/dist/taskpane.js"]) {
    if (!fs.existsSync(path.join(root, required))) throw new Error("RELEASE_INCOMPLETE");
  }
  const release = { id, root, directory, version: pkg.version, created: new Date().toISOString() };
  store.atomic(path.join(directory, ".release.json"), { ...release, files: inventory(directory) });
  verifyRelease(release); remember(dataDir, release); return release;
}
export function snapshotRelease(root, dataDir = store.home()) {
  const existing = Object.values(store.currentDeployment(dataDir).releases).find(release => path.resolve(release.root) === path.resolve(root));
  if (existing) { verifyRelease(existing); return existing; }
  const id = store.digest(crypto.randomUUID()), directory = path.join(dataDir, "releases", id);
  fs.mkdirSync(directory, { recursive: true });
  try {
    for (const name of ["bin", "dist", "gateway", "addin", "scripts", "node_modules", "package.json", "package-lock.json", "npm-shrinkwrap.json", "LICENSE"]) {
      const source = path.join(root, name);
      if (fs.existsSync(source)) fs.cpSync(source, path.join(directory, name), { recursive: true, verbatimSymlinks: true,
        filter: entry => { if (fs.lstatSync(entry).isSymbolicLink() && !store.inside(root, fs.realpathSync(entry))) throw new Error("SNAPSHOT_EXTERNAL_LINK"); return true; } });
    }
    return finalize(directory, directory, dataDir, id);
  } catch (error) { if (store.inside(path.join(dataDir, "releases"), directory)) fs.rmSync(directory, { recursive: true, force: true }); throw error; }
}
export async function prepareRelease(spec, { dataDir = store.home(), install = installPackage, run = runCommand, env = process.env } = {}) {
  if (!/^office-partner@\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(spec) && !(path.isAbsolute(spec) && /\.tgz$/i.test(spec))) throw new Error("EXACT_RELEASE_REQUIRED");
  const releases = path.join(dataDir, "releases"); fs.mkdirSync(releases, { recursive: true });
  const staging = fs.mkdtempSync(path.join(releases, ".install-")), id = store.digest(crypto.randomUUID());
  const directory = path.join(releases, id);
  try {
    store.atomic(path.join(staging, "package.json"), { name: "office-partner-deployment", version: "1.0.0", private: true });
    await install(staging, spec, { run, env });
    const candidate = path.join(staging, "node_modules", "office-partner");
    const pkg = store.readJson(path.join(candidate, "package.json"));
    if (spec.startsWith("office-partner@") && pkg.version !== spec.slice("office-partner@".length)) throw new Error("RELEASE_VERSION_MISMATCH");
    fs.renameSync(staging, directory);
    return finalize(directory, path.join(directory, "node_modules", "office-partner"), dataDir, id);
  } catch (error) {
    if (fs.existsSync(directory) && !store.currentDeployment(dataDir).releases[id]) fs.rmSync(directory, { recursive: true, force: true });
    throw error;
  } finally { if (fs.existsSync(staging)) fs.rmSync(staging, { recursive: true, force: true }); }
}
function journal(dataDir, patch) {
  const state = store.currentDeployment(dataDir); state.journal = { ...(state.journal ?? {}), ...patch, pid: process.pid, updated: new Date().toISOString() };
  store.atomic(path.join(dataDir, "deployment.json"), state);
}
export function restoreSelection(dataDir, before) {
  if (!before) return;
  store.changeState(dataDir, current => {
    for (const field of ["active", "previous", "selected", "disabled"]) {
      if (before[field] !== undefined) current[field] = before[field];
      else if (field === "disabled") delete current.disabled;
    }
  });
}
export async function switchRelease(next, options) {
  const { dataDir = store.home(), stop, start, drain, undrain = async () => {}, report = () => {}, profiles = {} } = options;
  const state = store.currentDeployment(dataDir), previous = state.releases[state.active];
  if (!previous) throw new Error("BASELINE_SNAPSHOT_REQUIRED");
  verifyRelease(next, true);
  let previousValid = true;
  try { verifyRelease(previous); } catch (error) { if (!options.allowDamagedCurrent) throw error; previousValid = false; }
  journal(dataDir, { phase: "prepared", from: previous.id, to: next.id, previous: state.previous, runtimeBefore: store.readState(dataDir) });
  let stopped = false;
  try {
    report("draining"); await drain();
    journal(dataDir, { phase: "stopping" }); stopped = true; await stop(previous);
    const deployment = store.currentDeployment(dataDir); deployment.active = next.id; deployment.previous = previous.id;
    const before = store.readState(dataDir);
    deployment.releases[previous.id].runtimeProfiles = { ...before.active };
    deployment.releases[previous.id].runtimeSelection = Object.fromEntries(["active", "previous", "selected", "disabled"].map(field => [field, before[field] ?? (field === "disabled" ? [] : undefined)]));
    deployment.releases[next.id].runtimeProfiles = options.selection?.active ?? { ...before.active, ...Object.fromEntries(Object.entries(profiles).map(([id, profile]) => [id, profile.id])) };
    deployment.held = options.hold === true; deployment.journal.phase = "activating"; store.atomic(path.join(dataDir, "deployment.json"), deployment);
    if (options.selection) restoreSelection(dataDir, options.selection);
    store.changeState(dataDir, runtimes => {
      for (const [provider, profile] of Object.entries(profiles)) {
        if (runtimes.active[provider] !== profile.id) runtimes.previous[provider] = runtimes.active[provider] ?? runtimes.legacy[provider];
        runtimes.active[provider] = profile.id;
        if (runtimes.pending?.[provider] === profile.id) delete runtimes.pending[provider];
      }
    });
    report("starting"); await start(next);
    journal(dataDir, { phase: "committed", runtimeBefore: null });
    return next;
  } catch (error) {
    report("recovering");
    if (!stopped) { await undrain(); journal(dataDir, { phase: "aborted", runtimeBefore: null }); throw error; }
    // Close partially started services before returning to the last complete installation.
    await stop(next).catch(() => {});
    const deployment = store.currentDeployment(dataDir), before = deployment.journal?.runtimeBefore;
    if (!previousValid) {
      // Keep the verified recovery target selected; the damaged source is not a rollback candidate.
      deployment.active = next.id; deployment.previous = null; deployment.held = true;
      deployment.journal = { ...deployment.journal, from: next.id, previous: null, runtimeBefore: store.readState(dataDir), phase: "recovery_failed" };
      store.atomic(path.join(dataDir, "deployment.json"), deployment);
      throw new Error("RECOVERY_RELEASE_START_FAILED_USE_RECOVERY");
    }
    deployment.active = previous.id; deployment.previous = deployment.journal?.previous ?? null; deployment.held = true;
    deployment.releases[next.id].failed = true; store.atomic(path.join(dataDir, "deployment.json"), deployment);
    restoreSelection(dataDir, before);
    try { await start(previous); journal(dataDir, { phase: "rolled_back", runtimeBefore: null }); }
    catch { journal(dataDir, { phase: "recovery_failed" }); throw new Error("ROLLBACK_START_FAILED_USE_RECOVERY"); }
    throw new Error("ACTIVATION_FAILED_PREVIOUS_RELEASE_RESTORED");
  }
}
export async function updateRelease(root, spec, options) {
  const dataDir = options.dataDir ?? store.home();
  return store.lock(dataDir, "deployment", async () => {
    let state = store.currentDeployment(dataDir);
    if (!state.active) {
      const baseline = snapshotRelease(root, dataDir); state = store.currentDeployment(dataDir); state.active = baseline.id;
      store.atomic(path.join(dataDir, "deployment.json"), state);
    }
    options.report?.("installing");
    const next = await prepareRelease(spec, { ...options, dataDir });
    const profiles = {};
    for (const provider of store.readState(dataDir).selected) profiles[provider] = await installRuntime(next.root, provider, { ...options, dataDir, adapterRoot: next.root, activate: false });
    return switchRelease(next, { ...options, dataDir, profiles });
  });
}
export async function rollbackRelease(version, options) {
  const dataDir = options.dataDir ?? store.home();
  return store.lock(dataDir, "deployment", async () => {
    const deployment = store.currentDeployment(dataDir);
    const next = version ? Object.values(deployment.releases).filter(release => release.version === version && !release.failed && release.id !== deployment.active).sort((a, b) => b.created.localeCompare(a.created))[0]
      : deployment.releases[deployment.previous];
    if (!next || next.id === deployment.active) throw new Error("ROLLBACK_RELEASE_UNAVAILABLE");
    const profiles = {};
    for (const [provider, id] of Object.entries(next.runtimeProfiles ?? {})) {
      const profile = store.readState(dataDir).profiles[id];
      if (profile) profiles[provider] = profile; else throw new Error("ROLLBACK_RUNTIME_UNAVAILABLE");
    }
    return switchRelease(next, { ...options, dataDir, profiles, selection: next.runtimeSelection, hold: true });
  });
}
export function recoverInterrupted(dataDir = store.home()) {
  const deployment = store.currentDeployment(dataDir), operation = deployment.journal;
  if (!operation || !["stopping", "activating", "recovery_failed"].includes(operation.phase)) return;
  if (Number.isInteger(operation.pid)) { try { process.kill(operation.pid, 0); return; } catch (error) { if (error.code !== "ESRCH") return; } }
  const previous = deployment.releases[operation.from]; if (!previous) throw new Error("RECOVERY_BASELINE_MISSING");
  verifyRelease(previous); deployment.active = previous.id; deployment.previous = operation.previous ?? null; deployment.held = true;
  restoreSelection(dataDir, operation.runtimeBefore);
  deployment.journal = { ...operation, phase: "recovered", runtimeBefore: null };
  store.atomic(path.join(dataDir, "deployment.json"), deployment);
}
