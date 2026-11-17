// Shared by the dependency-free bin tools and the compiled provider adapters.
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const manifest = require('./runtime-manifest.json');

const ids = Object.keys(manifest.providers);
const digest = value => crypto.createHash('sha256').update(value).digest('hex');
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
    fs.renameSync(temporary, file);
  }
  finally { try { fs.unlinkSync(temporary); } catch {} }
}

/** `selected` lists the providers whose CLI the product installed, `profiles` the installs, `active` the one each provider runs. */
const emptyState = () => ({ schemaVersion: 1, selected: [], profiles: {}, active: {}, disabled: [], failures: {} });
/** A missing or unreadable file is an empty state: the CLIs are installed again by the next setup or update. */
function readState(dataDir = home()) {
  try {
    const raw = readJson(path.join(dataDir, 'runtime-state.json'));
    const object = value => value && typeof value === 'object' && !Array.isArray(value) ? value : {};
    const list = value => Array.isArray(value) ? value.filter(id => ids.includes(id)) : [];
    return { ...emptyState(), selected: list(raw.selected), disabled: list(raw.disabled), active: object(raw.active), profiles: object(raw.profiles), failures: object(raw.failures) };
  } catch { return emptyState(); }
}
const bootTime = () => Math.round((Date.now() - os.uptime() * 1000) / 60000); // minutes: stable for one Windows session
const LOCK_MAX_AGE_MS = 2 * 60 * 60_000; // longer than any update, repair or setup runs
/**
 * A lock left behind is taken over when its process is gone, when Windows restarted since it was taken (its PID may now
 * belong to another program), when it is older than any operation runs, or when a crash left it empty.
 */
function staleLock(file, now = Date.now()) {
  let owner;
  try { owner = readJson(file); }
  catch { try { return now - fs.statSync(file).mtimeMs > 2 * 60_000; } catch { return true; } }
  if (!Number.isSafeInteger(owner.pid) || typeof owner.created !== 'number') return now - fs.statSync(file).mtimeMs > 2 * 60_000;
  if (now - owner.created > LOCK_MAX_AGE_MS) return true;
  if (typeof owner.boot === 'number' && Math.abs(owner.boot - bootTime()) > 1) return true;
  let alive = true;
  try { process.kill(owner.pid, 0); } catch (err) { alive = err.code !== 'ESRCH'; }
  return !alive && now - owner.created >= 30000;
}
function lock(dataDir, name, operation) {
  fs.mkdirSync(dataDir, { recursive: true });
  const file = path.join(dataDir, name + '.lock');
  let handle;
  try { handle = fs.openSync(file, 'wx', 0o600); }
  catch (error) {
    if (error.code !== 'EEXIST') throw error;
    if (!staleLock(file)) throw new Error('OPERATION_LOCKED');
    fs.unlinkSync(file); handle = fs.openSync(file, 'wx', 0o600);
  }
  fs.writeFileSync(handle, JSON.stringify({ pid: process.pid, created: Date.now(), boot: bootTime() }));
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

/** The install the provider runs, or null when it runs the system CLI. */
function selectProfile(dataDir, provider) {
  const state = readState(dataDir);
  if (state.disabled.includes(provider)) throw new Error('PROVIDER_DISABLED');
  const id = state.active[provider];
  if (!id) return null;
  const profile = state.profiles[id];
  if (!profile) throw new Error('RUNTIME_PROFILE_MISSING');
  if (profile.platform !== process.platform || profile.arch !== process.arch) throw new Error('RUNTIME_PLATFORM_MISMATCH');
  if (!fs.existsSync(profile.command)) throw new Error('RUNTIME_MISSING');
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
/**
 * The admin policy's update settings (the whole policy is protocol/src/policy.ts; bin/ runs on its own, so it reads the
 * same file here). No file: updates allowed. A file that cannot be read or parsed: updates off, as everything risky is.
 */
function policyFile(env = process.env) {
  return env.OFFICE_PARTNER_POLICY_FILE || path.join(env.ProgramData || env.PROGRAMDATA || 'C:\\ProgramData', 'Office Partner', 'policy.json');
}
function readPolicy(env = process.env) {
  let text;
  try { text = fs.readFileSync(policyFile(env), 'utf8').replace(/^\uFEFF/, ''); }
  catch (error) { return error.code === 'ENOENT' || error.code === 'ENOTDIR' ? { autoUpdate: true } : { autoUpdate: false, error: 'POLICY_UNREADABLE' }; }
  let raw;
  try { raw = JSON.parse(text); } catch { return { autoUpdate: false, error: 'POLICY_INVALID' }; }
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)
    || raw.autoUpdate !== undefined && typeof raw.autoUpdate !== 'boolean'
    || raw.pinVersion !== undefined && !/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(raw.pinVersion)
    || raw.registry !== undefined && !/^https?:\/\/\S+$/.test(raw.registry)) return { autoUpdate: false, error: 'POLICY_INVALID' };
  return { autoUpdate: raw.autoUpdate !== false, pinVersion: raw.pinVersion, registry: raw.registry };
}
module.exports = { readPolicy, staleLock, manifest, ids, digest, home, inside, readJson, atomic, renameFolder, emptyState, readState, changeState, lock, selectProfile, childEnvironment };
