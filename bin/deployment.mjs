import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import store from "./runtime-store.cjs";
import { dropUnused, installPackage, installRuntime, npmEntrypoint, verifyRuntime } from "./runtime-manager.mjs";
import { runCommand } from "./doctor-probes.mjs";

const fileHash = file => store.fileDigest(file);
/** `known`: hashes of files that are the same file as one already recorded (hard-linked from the previous release), so they are not read again. */
function inventory(directory, known = {}) {
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
      else files[relative] = known[relative] ?? { sha256: fileHash(file) };
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
function finalize(directory, root, dataDir, id, known = {}) {
  const pkg = store.readJson(path.join(root, "package.json"));
  if (pkg.name !== "office-partner" || !/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(pkg.version)) throw new Error("RELEASE_PACKAGE_INVALID");
  for (const required of ["bin/office-partner.mjs", "bin/runtime-store.cjs", "bin/runtime-manifest.json", "dist/bridge/src/index.js", "gateway/dist/server.mjs", "frontend/dist/taskpane.js"]) {
    if (!fs.existsSync(path.join(root, required))) throw new Error("RELEASE_INCOMPLETE");
  }
  const release = { id, root, directory, version: pkg.version, created: new Date().toISOString() };
  // `modules`: the native-module ABI the dependencies were installed for; a release may reuse another's dependencies only on the same one.
  store.atomic(path.join(directory, ".release.json"), { ...release, modules: process.versions.modules, files: inventory(directory, known) });
  verifyRelease(release); remember(dataDir, release); return release;
}
export function snapshotRelease(root, dataDir = store.home()) {
  const existing = Object.values(store.currentDeployment(dataDir).releases).find(release => path.resolve(release.root) === path.resolve(root));
  if (existing) { verifyRelease(existing); return existing; }
  const id = store.digest(crypto.randomUUID()), directory = path.join(dataDir, "releases", id);
  fs.mkdirSync(directory, { recursive: true });
  try {
    for (const name of ["bin", "dist", "gateway", "frontend", "scripts", "node_modules", "package.json", "package-lock.json", "npm-shrinkwrap.json", "LICENSE"]) {
      const source = path.join(root, name);
      if (fs.existsSync(source)) fs.cpSync(source, path.join(directory, name), { recursive: true, verbatimSymlinks: true,
        filter: entry => { if (fs.lstatSync(entry).isSymbolicLink() && !store.inside(root, fs.realpathSync(entry))) throw new Error("SNAPSHOT_EXTERNAL_LINK"); return true; } });
    }
    return finalize(directory, directory, dataDir, id);
  } catch (error) { if (store.inside(path.join(dataDir, "releases"), directory)) fs.rmSync(directory, { recursive: true, force: true }); throw error; }
}
function linkTree(source, target) {
  fs.mkdirSync(target, { recursive: true });
  for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
    const from = path.join(source, entry.name), to = path.join(target, entry.name);
    if (entry.isDirectory()) linkTree(from, to);
    else if (entry.isFile()) fs.linkSync(from, to);
    else throw new Error("LINK_UNSUPPORTED");
  }
}
/**
 * Builds `staging` as a release that shares the dependency files of the running one: the new product is unpacked from its tarball and every
 * other file of `node_modules` is a hard link, so nothing is downloaded or copied and the old release can be deleted without touching it.
 * Returns the recorded hashes of the shared files, or null when the dependencies are not provably the same (then npm installs them).
 */
async function reuseDependencies(spec, staging, dataDir, { run, env }) {
  const state = store.currentDeployment(dataDir), active = state.releases[state.active];
  if (!active || path.resolve(active.root) !== path.join(active.directory, "node_modules", "office-partner")) return null;
  const record = store.readJson(path.join(active.directory, ".release.json"), 64 * 1024 * 1024);
  if (record.modules !== process.versions.modules) return null;
  let tarball = spec;
  if (!/\.tgz$/i.test(spec)) {
    const packed = path.join(staging, ".packed");
    const result = await run(process.execPath, [npmEntrypoint(env), "pack", spec, "--pack-destination", packed, "--prefer-offline", "--no-audit", "--no-fund"], { env, timeoutMs: 5 * 60_000 });
    const names = result.code === 0 && fs.existsSync(packed) ? fs.readdirSync(packed).filter(name => name.endsWith(".tgz")) : [];
    if (names.length !== 1) return null;
    tarball = path.join(packed, names[0]);
  }
  const unpacked = path.join(staging, ".unpacked");
  fs.mkdirSync(unpacked, { recursive: true });
  const tar = process.platform === "win32" ? path.join(process.env.SystemRoot ?? "C:\\Windows", "System32", "tar.exe") : "tar";
  if ((await run(tar, ["-xzf", tarball, "-C", unpacked], { env, timeoutMs: 2 * 60_000 })).code !== 0) return null;
  const product = path.join(unpacked, "package"), wanted = path.join(product, "npm-shrinkwrap.json"), current = path.join(active.root, "npm-shrinkwrap.json");
  if (!fs.existsSync(wanted) || !fs.existsSync(current) || fileHash(wanted) !== fileHash(current)) return null;
  const modules = path.join(staging, "node_modules");
  fs.mkdirSync(modules, { recursive: true });
  for (const name of fs.readdirSync(path.join(active.directory, "node_modules"))) {
    if (name === "office-partner") continue;
    const from = path.join(active.directory, "node_modules", name);
    if (fs.statSync(from).isDirectory()) linkTree(from, path.join(modules, name)); else fs.linkSync(from, path.join(modules, name));
  }
  const nested = path.join(active.root, "node_modules");
  if (fs.existsSync(nested)) linkTree(nested, path.join(product, "node_modules"));
  fs.renameSync(product, path.join(modules, "office-partner"));
  fs.rmSync(unpacked, { recursive: true, force: true });
  fs.rmSync(path.join(staging, ".packed"), { recursive: true, force: true });
  const shared = {};
  for (const [relative, entry] of Object.entries(record.files)) if (relative.startsWith("node_modules" + path.sep) && !relative.startsWith(path.join("node_modules", "office-partner") + path.sep)) shared[relative] = entry;
  // Files that were nested in the old product folder are shared too, under the same relative path.
  for (const [relative, entry] of Object.entries(record.files)) if (relative.startsWith(path.join("node_modules", "office-partner", "node_modules") + path.sep)) shared[relative] = entry;
  return shared;
}
export async function prepareRelease(spec,{ dataDir = store.home(), install = installPackage, run = runCommand, env = process.env } = {}) {
  if (!/^office-partner@\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(spec) && !(path.isAbsolute(spec) && /\.tgz$/i.test(spec))) throw new Error("EXACT_RELEASE_REQUIRED");
  const releases = path.join(dataDir, "releases"); fs.mkdirSync(releases, { recursive: true });
  const staging = fs.mkdtempSync(path.join(releases, ".install-")), id = store.digest(crypto.randomUUID());
  const directory = path.join(releases, id);
  try {
    // Dependencies rarely change between releases, and installing them is most of an update: when they are the same, the new release
    // is the new product files next to hard links of the dependency files already on disk.
    let known = (install === installPackage) ? await reuseDependencies(spec, staging, dataDir, { run, env }).catch(() => null) : null;
    if (!known) {
      for (const name of fs.readdirSync(staging)) fs.rmSync(path.join(staging, name), { recursive: true, force: true });
      store.atomic(path.join(staging, "package.json"), { name: "office-partner-deployment", version: "1.0.0", private: true });
      await install(staging, spec, { run, env });
      known = {};
    }
    const candidate = path.join(staging, "node_modules", "office-partner");
    const pkg = store.readJson(path.join(candidate, "package.json"));
    if (spec.startsWith("office-partner@") && pkg.version !== spec.slice("office-partner@".length)) throw new Error("RELEASE_VERSION_MISMATCH");
    store.renameFolder(staging, directory);
    return finalize(directory, path.join(directory, "node_modules", "office-partner"), dataDir, id, known);
  } catch (error) {
    if (fs.existsSync(directory) && !store.currentDeployment(dataDir).releases[id]) fs.rmSync(directory, { recursive: true, force: true });
    throw error;
  } finally { if (fs.existsSync(staging)) fs.rmSync(staging, { recursive: true, force: true }); }
}
/** The runtimes a new release will use must be intact and for this machine before anything is stopped. */
function verifyProfiles(profiles, dataDir) {
  for (const profile of Object.values(profiles)) {
    if (profile.platform !== process.platform || profile.arch !== process.arch) throw new Error("RUNTIME_PLATFORM_MISMATCH");
    verifyRuntime(profile, dataDir);
  }
}
/**
 * Keep only the running release, and only the CLI installs a provider runs, on disk. Called once a switch has settled: committed, or rolled
 * back to the old release. The old release existed to undo a failed switch; an older version is installed again from npm when someone wants it.
 *
 * Deleting a folder takes about half a minute on Windows (every file is scanned), so the update must not wait for it:
 * each folder is renamed to `.trash-*` at once and removed in the background. `pending` settles when that is finished.
 * Cleanup trouble (a file held open by antivirus, for example) never fails an update; the next call sweeps what is left.
 */
export function retireOldReleases(dataDir = store.home()) {
  const result = { releases: [], runtimes: [], pending: Promise.resolve() };
  try {
    const deployment = store.currentDeployment(dataDir), active = deployment.releases[deployment.active];
    if (!active) return result;
    const retired = Object.values(deployment.releases).filter(release => release.id !== active.id);
    for (const release of retired) delete deployment.releases[release.id];
    store.atomic(path.join(dataDir, "deployment.json"), deployment);
    store.changeState(dataDir, dropUnused);
    result.releases = retired.map(release => release.id);

    const releases = path.join(dataDir, "releases"), runtimes = path.join(dataDir, "runtimes");
    const names = directory => fs.existsSync(directory) ? fs.readdirSync(directory) : [];
    const doomed = retired.map(release => release.directory);
    // Folders that no longer have a record (an earlier cleanup could not delete them) and trash left by an interrupted cleanup.
    for (const name of names(releases)) if (/^[a-f0-9]{64}$/.test(name) && !deployment.releases[name] || name.startsWith(".trash-")) doomed.push(path.join(releases, name));
    // CLI installs no provider runs any more: the one a provider moved away from, one that was only tried, one found damaged.
    const inUse = new Set(Object.values(store.readState(dataDir).active));
    for (const name of names(runtimes)) {
      if (!name.startsWith(".trash-") && !(/^[a-f0-9]{64}/.test(name) && (!inUse.has(name) || name.includes(".damaged-")))) continue;
      doomed.push(path.join(runtimes, name));
      if (!name.startsWith(".trash-")) result.runtimes.push(name);
    }
    const removals = [];
    for (const directory of new Set(doomed)) {
      if (!fs.existsSync(directory) || !store.inside(dataDir, directory)) continue;
      let target = directory;
      if (!path.basename(directory).startsWith(".trash-")) {
        const trash = path.join(path.dirname(directory), ".trash-" + crypto.randomUUID());
        try { fs.renameSync(directory, trash); target = trash; } catch { /* held open: delete in place, or sweep next time */ }
      }
      removals.push(fs.promises.rm(target, { recursive: true, force: true, maxRetries: 3, retryDelay: 200 }).catch(() => {}));
    }
    result.pending = Promise.all(removals).then(() => {});
  } catch { /* A failed cleanup must not turn a good update into a failed one. */ }
  return result;
}
/**
 * Stop the services, make `next` the active release and start it. A failed start is reported as it is: nothing goes back to the
 * older release by itself. To go back to an older version, run `office-partner rollback <version>`, which installs it from npm.
 */
export async function switchRelease(next, options) {
  const { dataDir = store.home(), stop, start, drain, undrain = async () => {}, report = () => {}, profiles = {} } = options;
  const state = store.currentDeployment(dataDir), previous = state.releases[state.active];
  if (!previous) throw new Error("BASELINE_SNAPSHOT_REQUIRED");
  verifyRelease(next, true);
  verifyProfiles(profiles, dataDir);
  let stopped = false;
  try {
    report("draining"); await drain();
    stopped = true; await stop(previous);
    const deployment = store.currentDeployment(dataDir); deployment.active = next.id;
    store.atomic(path.join(dataDir, "deployment.json"), deployment);
    store.changeState(dataDir, runtimes => {
      for (const [provider, profile] of Object.entries(profiles)) runtimes.active[provider] = profile.id;
      dropUnused(runtimes);
    });
    report("starting"); await start(next);
    // The older release is no longer needed once the new one runs.
    retireOldReleases(dataDir);
    return next;
  } catch (error) {
    if (stopped) await stop(next).catch(() => {}); // close what only started halfway
    await undrain();
    throw error;
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
    for (const provider of store.readState(dataDir).selected) {
      try { profiles[provider] = await installRuntime(next.root, provider, { ...options, dataDir, activate: false }); }
      catch (error) {
        // A provider whose new CLI cannot be installed or fails its checks must not stop everyone else from updating: its current CLI
        // keeps running, and the reason is recorded for `doctor`.
        options.notice?.(`${provider}: the new CLI was not installed (${error instanceof Error ? error.message : error}), so the current one stays.`);
      }
    }
    return switchRelease(next, { ...options, dataDir, profiles });
  });
}
