// Shared by the dependency-free recovery CLI and the compiled provider adapters.
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const manifest = require('./runtime-manifest.json');

const ids = Object.keys(manifest.providers);
const digest = value => crypto.createHash('sha256').update(value).digest('hex');
function fileDigest(file) {
  const hash = crypto.createHash('sha256'), handle = fs.openSync(file, 'r'), buffer = Buffer.allocUnsafe(1024 * 1024);
  try { let count; while ((count = fs.readSync(handle, buffer, 0, buffer.length, null)) > 0) hash.update(buffer.subarray(0, count)); return hash.digest('hex'); }
  finally { fs.closeSync(handle); }
}
const home = (env = process.env) => env.OFFICE_PARTNER_DATA_DIR || path.join(env.LOCALAPPDATA || os.homedir(), 'Excel Partner');
function inside(root, candidate) {
  const relative = path.relative(path.resolve(root), path.resolve(candidate));
  return relative === '' || !relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative);
}
function readJson(file, limit = 8 * 1024 * 1024) {
  if (fs.statSync(file).size > limit) throw new Error('STATE_TOO_LARGE');
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}
/**
 * Rename a freshly written folder. On Windows the virus scanner holds new files open for a while, and the rename then fails with
 * EPERM/EBUSY/EACCES even though nothing is wrong, so it is tried again for up to about 40 seconds before the error is let through.
 */
function renameFolder(from, to) {
  const wait = new Int32Array(new SharedArrayBuffer(4));
  for (let attempt = 0; ; attempt++) {
    try { fs.renameSync(from, to); return; }
    catch (error) {
      if (!['EPERM', 'EBUSY', 'EACCES'].includes(error.code) || process.platform !== 'win32' || attempt >= 40) throw error;
      Atomics.wait(wait, 0, 0, 500 + attempt * 50);
    }
  }
}
function atomic(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const temporary = file + '.' + crypto.randomUUID() + '.tmp';
  try {
    const handle = fs.openSync(temporary, 'wx', 0o600);
    try { fs.writeFileSync(handle, JSON.stringify(data, null, 2) + '\n'); fs.fsyncSync(handle); } finally { fs.closeSync(handle); }
    if (['runtime-state.json', 'deployment.json'].includes(path.basename(file)) && fs.existsSync(file)) fs.copyFileSync(file, file + '.previous');
    fs.renameSync(temporary, file);
  }
  finally { try { fs.unlinkSync(temporary); } catch {} }
}

/** `selected` lists the providers whose CLI the product installed, `profiles` the installs, `active` the one each provider runs. */
const emptyState = () => ({ schemaVersion: 1, selected: [], profiles: {}, active: {}, owners: {}, disabled: [], failures: {} });
function readState(dataDir = home(), suffix = '') {
  const file = path.join(dataDir, 'runtime-state.json' + suffix);
  if (!fs.existsSync(file)) return emptyState();
  const raw = readJson(file);
  if (raw.schemaVersion !== 1 || !Array.isArray(raw.selected) || !raw.selected.every(id => ids.includes(id))) throw new Error('RUNTIME_STATE_INVALID');
  for (const field of ['profiles', 'active', 'owners']) {
    if (!raw[field] || typeof raw[field] !== 'object' || Array.isArray(raw[field])) throw new Error('RUNTIME_STATE_INVALID');
  }
  // Older versions also kept previous CLIs, one CLI per saved chat and copies of the system CLI. Only the install in use matters now.
  const state = { ...emptyState(), selected: raw.selected, active: raw.active, owners: raw.owners, disabled: raw.disabled ?? [], failures: raw.failures ?? {},
    profiles: Object.fromEntries(Object.entries(raw.profiles).filter(([, profile]) => profile.mode !== 'external')) };
  for (const [id, profile] of Object.entries(state.profiles)) {
    if (profile.id !== id || !ids.includes(profile.provider) || typeof profile.command !== 'string' || !path.isAbsolute(profile.command)) throw new Error('RUNTIME_PROFILE_INVALID');
    if (!inside(path.join(dataDir, 'runtimes'), profile.command)) throw new Error('RUNTIME_PATH_INVALID');
    if (profile.stateHome && (!path.isAbsolute(profile.stateHome) || !inside(path.join(dataDir, 'provider-data'), profile.stateHome))) throw new Error('RUNTIME_STATE_HOME_INVALID');
  }
  return state;
}
function lock(dataDir, name, operation) {
  fs.mkdirSync(dataDir, { recursive: true });
  const file = path.join(dataDir, name + '.lock');
  let handle;
  try { handle = fs.openSync(file, 'wx', 0o600); }
  catch (error) {
    if (error.code !== 'EEXIST') throw error;
    let owner;
    try { owner = readJson(file); } catch { throw new Error('OPERATION_LOCKED'); }
    let alive = true;
    try { process.kill(owner.pid, 0); } catch (err) { alive = err.code !== 'ESRCH'; }
    if (alive || Date.now() - owner.created < 30000) throw new Error('OPERATION_LOCKED');
    fs.unlinkSync(file); handle = fs.openSync(file, 'wx', 0o600);
  }
  fs.writeFileSync(handle, JSON.stringify({ pid: process.pid, created: Date.now() }));
  const release = () => { fs.closeSync(handle); try { fs.unlinkSync(file); } catch {} };
  try {
    const result = operation();
    if (result && typeof result.then === 'function') return result.finally(release);
    release(); return result;
  } catch (error) { release(); throw error; }
}
function changeState(dataDir, change) {
  return lock(dataDir, 'runtime-state', () => { const state = readState(dataDir); const result = change(state); atomic(path.join(dataDir, 'runtime-state.json'), state); return result; });
}

/** True while the CLI file is the one that was installed. A full read of a large file is only repeated when its size or time changes. */
const checked = new Map();
function unchanged(profile) {
  const stat = fs.statSync(profile.command), signature = `${stat.size}:${stat.mtimeMs}:${profile.fingerprint}`;
  if (checked.get(profile.command) === signature) return true;
  if (fileDigest(profile.command) !== profile.fingerprint) return false;
  checked.set(profile.command, signature);
  return true;
}
/** The install the provider runs, or null when it runs the system CLI. */
function selectProfile(dataDir, provider) {
  const state = readState(dataDir);
  if (state.disabled.includes(provider)) throw new Error('PROVIDER_DISABLED');
  const id = state.active[provider];
  if (!id) return null;
  const profile = state.profiles[id];
  if (!profile) throw new Error('RUNTIME_PROFILE_MISSING');
  if (profile.platform !== process.platform || profile.arch !== process.arch) throw new Error('RUNTIME_PLATFORM_MISMATCH');
  if (!fs.existsSync(profile.command) || !unchanged(profile)) throw new Error('RUNTIME_MISSING');
  return profile;
}
/** The environment a provider's CLI runs in: updates off, its own state folder, and the install the product chose. */
function childEnvironment(env, profile) {
  const result = { ...env, ...manifest.autoUpdate };
  if (!profile) return result;
  const spec = manifest.providers[profile.provider].env ?? {};
  if (profile.command) result[profile.provider.toUpperCase() + '_BIN'] = profile.command;
  if (profile.stateHome) for (const [name, relative] of Object.entries(spec.state ?? {})) result[name] = path.join(profile.stateHome, relative);
  for (const name of spec.unset ?? []) delete result[name];
  return result;
}
function currentDeployment(dataDir = home(), suffix = '') {
  const file = path.join(dataDir, 'deployment.json' + suffix);
  if (!fs.existsSync(file)) return { schemaVersion: 1, active: null, releases: {} };
  const data = readJson(file);
  if (data.schemaVersion !== 1 || !data.releases || typeof data.releases !== 'object') throw new Error('DEPLOYMENT_STATE_INVALID');
  for (const [id, release] of Object.entries(data.releases)) {
    if (!/^[a-f0-9]{64}$/.test(id) || release.id !== id || typeof release.directory !== 'string' || !path.isAbsolute(release.directory)
      || !inside(path.join(dataDir, 'releases'), release.directory) || !inside(release.directory, release.root)) throw new Error('RELEASE_PATH_INVALID');
  }
  return data;
}
module.exports = { manifest, ids, digest, fileDigest, home, inside, readJson, atomic, renameFolder, emptyState, readState, changeState, lock, selectProfile, childEnvironment, currentDeployment };
