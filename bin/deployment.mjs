import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import store from "./runtime-store.cjs";
import { installPackage, installRuntime, verifyRuntime } from "./runtime-manager.mjs";
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
/** The runtimes a new release will use must be intact and for this machine before anything is stopped. */
function verifyProfiles(profiles, dataDir) {
  for (const profile of Object.values(profiles)) {
    if (!profile) throw new Error("RUNTIME_UNAVAILABLE");
    if (profile.mode === "managed") {
      if (profile.platform !== process.platform || profile.arch !== process.arch) throw new Error("RUNTIME_PLATFORM_MISMATCH");
      verifyRuntime(profile, dataDir);
    } else if (!fs.existsSync(profile.command) || profile.fingerprint && store.fileDigest(profile.command) !== profile.fingerprint) throw new Error("RUNTIME_UNAVAILABLE");
  }
}
/**
 * Keep only the running release on disk. Called once a switch has settled: committed, or rolled back to the old release.
 * The old release existed to undo a failed switch; an older version is installed again from npm when someone wants it.
 *
 * Saved chats are looked after: a runtime profile that only exists because a retired release created it is dropped when it
 * duplicates the current CLI (same version and data folder), and its chats move to the current profile. An older CLI that
 * chats may still need to resume with is kept, now served by the current release's adapter.
 *
 * Deleting a release folder takes about half a minute on Windows (every file is scanned), so the update must not wait for it:
 * each folder is renamed to `.trash-*` at once and removed in the background. `pending` settles when that is finished.
 * Cleanup trouble (a file held open by antivirus, for example) never fails an update; the next call sweeps what is left.
 */
export function retireOldReleases(dataDir = store.home()) {
  const result = { releases: [], profiles: [], pending: Promise.resolve() };
  try {
    const deployment = store.currentDeployment(dataDir), active = deployment.releases[deployment.active];
    if (!active) return result;
    const retired = Object.values(deployment.releases).filter(release => release.id !== active.id);
    const pinned = root => typeof root === "string" && retired.some(release => store.inside(release.directory, root));
    const dropped = [];
    if (retired.length) store.changeState(dataDir, state => {
      for (const [id, profile] of Object.entries(state.profiles)) {
        if (!pinned(profile.adapterRoot)) continue;
        const provider = profile.provider, current = state.profiles[state.active[provider]];
        const inUse = state.active[provider] === id || state.legacy[provider] === id;
        const bound = Object.entries(state.conversations).filter(([, profileId]) => profileId === id);
        const wanted = inUse || bound.length > 0 || state.previous[provider] === id;
        const duplicate = !inUse && current && current.id !== id && current.version === profile.version && current.stateHome === profile.stateHome;
        if (state.pending?.[provider] === id || !wanted || duplicate) {
          for (const [key] of bound) { if (duplicate) state.conversations[key] = current.id; else delete state.conversations[key]; }
          if (state.previous[provider] === id) delete state.previous[provider];
          if (state.pending?.[provider] === id) delete state.pending[provider];
          delete state.profiles[id]; dropped.push(profile);
        } else profile.adapterRoot = active.root; // an older CLI some chats still resume with
      }
    });
    for (const release of retired) delete deployment.releases[release.id];
    store.atomic(path.join(dataDir, "deployment.json"), deployment);
    result.releases = retired.map(release => release.id); result.profiles = dropped.map(profile => profile.id);

    const releases = path.join(dataDir, "releases"), runtimes = path.join(dataDir, "runtimes");
    const doomed = retired.map(release => release.directory);
    for (const profile of dropped) if (profile.mode === "managed") doomed.push(path.join(runtimes, profile.id));
    // Folders that no longer have a record (an earlier cleanup could not delete them) and trash left by an interrupted cleanup.
    for (const name of fs.existsSync(releases) ? fs.readdirSync(releases) : []) {
      if (/^[a-f0-9]{64}$/.test(name) && !deployment.releases[name] || name.startsWith(".trash-")) doomed.push(path.join(releases, name));
    }
    for (const name of fs.existsSync(runtimes) ? fs.readdirSync(runtimes) : []) if (name.startsWith(".trash-")) doomed.push(path.join(runtimes, name));
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
      for (const [provider, profile] of Object.entries(profiles)) {
        if (runtimes.active[provider] !== profile.id) runtimes.previous[provider] = runtimes.active[provider] ?? runtimes.legacy[provider];
        runtimes.active[provider] = profile.id;
        if (runtimes.pending?.[provider] === profile.id) delete runtimes.pending[provider];
      }
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
    for (const provider of store.readState(dataDir).selected) profiles[provider] = await installRuntime(next.root, provider, { ...options, dataDir, adapterRoot: next.root, activate: false });
    return switchRelease(next, { ...options, dataDir, profiles });
  });
}
