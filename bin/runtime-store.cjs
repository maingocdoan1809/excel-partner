// Shared by the dependency-free recovery CLI and the compiled provider adapters.
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');

const ids = ['claude', 'copilot', 'codex', 'opencode'];
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
const empty = () => ({ schemaVersion: 1, selected: [], profiles: {}, active: {}, previous: {}, legacy: {}, conversations: {}, owners: {} });
function readState(dataDir = home(), suffix = '') {
  const file = path.join(dataDir, 'runtime-state.json' + suffix);
  if (!fs.existsSync(file)) return empty();
  const state = readJson(file);
  if (state.schemaVersion !== 1 || !Array.isArray(state.selected) || !state.selected.every(id => ids.includes(id))) throw new Error('RUNTIME_STATE_INVALID');
  for (const field of ['profiles', 'active', 'previous', 'legacy', 'conversations', 'owners']) {
    if (!state[field] || typeof state[field] !== 'object' || Array.isArray(state[field])) throw new Error('RUNTIME_STATE_INVALID');
  }
  for (const [id, profile] of Object.entries(state.profiles)) {
    if (!/^[a-f0-9]{64}$/.test(id) || id !== profile.id || !ids.includes(profile.provider)
      || !['managed', 'external'].includes(profile.mode) || typeof profile.command !== 'string' || !path.isAbsolute(profile.command)
      || typeof profile.adapterRoot !== 'string' || !path.isAbsolute(profile.adapterRoot)) throw new Error('RUNTIME_PROFILE_INVALID');
    if (profile.mode === 'managed' && (!inside(path.join(dataDir, 'runtimes'), profile.command)
      || !inside(path.join(dataDir, 'releases'), profile.adapterRoot))) throw new Error('RUNTIME_PATH_INVALID');
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
const conversationKey = (provider, nativeId) => provider + ':' + digest(nativeId);
function bindConversation(dataDir, provider, nativeId, profileId) {
  if (!nativeId || !profileId) return;
  if (readState(dataDir).conversations[conversationKey(provider, nativeId)] === profileId) return;
  changeState(dataDir, state => {
    const key = conversationKey(provider, nativeId), previous = state.conversations[key];
    if (previous && previous !== profileId) throw new Error('CONVERSATION_PROFILE_CONFLICT');
    if (Object.keys(state.conversations).length >= 10000 && !previous) throw new Error('CONVERSATION_PROFILE_LIMIT');
    if (!state.profiles[profileId]) throw new Error('RUNTIME_PROFILE_MISSING');
    state.conversations[key] = profileId;
  });
}
function selectProfile(dataDir, provider, nativeId) {
  const state = readState(dataDir);
  if (!nativeId && state.disabled?.includes(provider)) throw new Error('PROVIDER_DISABLED');
  const saved = nativeId ? state.conversations[conversationKey(provider, nativeId)] : undefined;
  // An old conversation with no recorded profile stays on its captured external CLI.
  const id = saved || (nativeId ? state.legacy[provider] : state.active[provider] || state.legacy[provider]);
  if (!id && nativeId && state.active[provider]) throw new Error('LEGACY_RESUME_PROFILE_UNKNOWN');
  if (!id) return null;
  const profile = state.profiles[id];
  if (!profile) throw new Error('RUNTIME_PROFILE_MISSING');
  if (profile.mode === 'managed' && (profile.platform !== process.platform || profile.arch !== process.arch)) throw new Error('RUNTIME_PLATFORM_MISMATCH');
  if (!fs.existsSync(profile.command) || profile.fingerprint && fileDigest(profile.command) !== profile.fingerprint) {
    if (nativeId) throw new Error('RESUME_RUNTIME_MISSING');
    const fallback = fallbackProfile(dataDir, provider, id);
    if (!fallback) throw new Error('RUNTIME_MISSING');
    return { ...fallback, fallbackFrom: id };
  }
  return profile;
}
function fallbackProfile(dataDir, provider, excluded) {
  const state = readState(dataDir);
  for (const id of [state.previous[provider], state.legacy[provider]]) {
    const profile = state.profiles[id];
    if (profile && id !== excluded && fs.existsSync(profile.command) && (!profile.fingerprint || fileDigest(profile.command) === profile.fingerprint)) return profile;
  }
  return null;
}
function childEnvironment(env, profile) {
  const result = { ...env, COPILOT_AUTO_UPDATE: 'false', DISABLE_AUTOUPDATER: '1', DISABLE_UPDATES: '1', OPENCODE_DISABLE_AUTOUPDATE: 'true' };
  if (profile) {
    if (profile.command) result[profile.provider.toUpperCase() + '_BIN'] = profile.command;
    if (profile.mode === 'managed' && profile.stateHome) {
      if (profile.provider === 'copilot') { result.COPILOT_HOME = profile.stateHome; result.COPILOT_CACHE_HOME = path.join(profile.stateHome, 'cache'); }
      if (profile.provider === 'claude') result.CLAUDE_CONFIG_DIR = profile.stateHome;
      if (profile.provider === 'codex') result.CODEX_HOME = profile.stateHome;
      if (profile.provider === 'opencode') {
        delete result.OPENCODE_BIN_PATH;
        for (const [key, directory] of Object.entries({ XDG_CONFIG_HOME: 'config', XDG_DATA_HOME: 'data', XDG_STATE_HOME: 'state', XDG_CACHE_HOME: 'cache' })) result[key] = path.join(profile.stateHome, directory);
        result.OPENCODE_CONFIG_DIR = path.join(profile.stateHome, 'config', 'opencode');
        delete result.OPENCODE_DB;
      }
    }
    if (profile.provider === 'copilot') delete result.COPILOT_CLI_PATH;
  }
  return result;
}
function currentDeployment(dataDir = home(), suffix = '') {
  const file = path.join(dataDir, 'deployment.json' + suffix);
  if (!fs.existsSync(file)) return { schemaVersion: 1, active: null, previous: null, releases: {}, journal: null };
  const data = readJson(file);
  if (data.schemaVersion !== 1 || !data.releases || typeof data.releases !== 'object') throw new Error('DEPLOYMENT_STATE_INVALID');
  for (const [id, release] of Object.entries(data.releases)) {
    if (!/^[a-f0-9]{64}$/.test(id) || release.id !== id || typeof release.directory !== 'string' || !path.isAbsolute(release.directory)
      || !inside(path.join(dataDir, 'releases'), release.directory) || !inside(release.directory, release.root)) throw new Error('RELEASE_PATH_INVALID');
  }
  return data;
}
module.exports = { ids, digest, fileDigest, home, inside, readJson, atomic, readState, changeState, lock, bindConversation, selectProfile, fallbackProfile, childEnvironment, currentDeployment, conversationKey };
