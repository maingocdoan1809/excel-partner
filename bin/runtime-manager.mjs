import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import store from "./runtime-store.cjs";
import { resolveCommands } from "./doctor-checks.mjs";
import { runCommand } from "./doctor-probes.mjs";

const packages = { claude: "@anthropic-ai/claude-code", copilot: "@github/copilot", codex: "@openai/codex", opencode: "opencode-ai" };
export function readManifest(root) {
  const manifest = store.readJson(path.join(root, "bin", "runtime-manifest.json"));
  if (manifest.schemaVersion !== 1 || !manifest.providers) throw new Error("RUNTIME_MANIFEST_INVALID");
  for (const id of store.ids) {
    const item = manifest.providers[id];
    if (!item || item.package !== packages[id] || item.bin !== id || !/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(item.version)
      || item.sdkVersion && !/^\d+\.\d+\.\d+$/.test(item.sdkVersion)) throw new Error("RUNTIME_MANIFEST_INVALID");
  }
  return manifest;
}
export function npmEntrypoint(env = process.env) {
  const candidates = [env.npm_execpath, path.join(path.dirname(process.execPath), "node_modules", "npm", "bin", "npm-cli.js")];
  for (const shim of resolveCommands("npm", env)) candidates.push(shim, path.join(path.dirname(shim), "node_modules", "npm", "bin", "npm-cli.js"), path.join(path.dirname(shim), "..", "lib", "node_modules", "npm", "bin", "npm-cli.js"));
  for (const candidate of candidates) if (candidate && path.basename(candidate) === "npm-cli.js" && fs.existsSync(candidate)) return fs.realpathSync(candidate);
  throw new Error("NPM_NOT_FOUND");
}
export async function installPackage(prefix, spec, { env = process.env, run = runCommand } = {}) {
  // No shell. npm always receives arguments as separate strings, including paths with spaces.
  const installEnv = { ...env };
  for (const key of Object.keys(installEnv)) if (key.toLowerCase() === "npm_config_engine_strict") delete installEnv[key];
  installEnv.NPM_CONFIG_ENGINE_STRICT = "true";
  const result = await run(process.execPath, [npmEntrypoint(env), "install", "--global=false", "--prefix", prefix, "--save-exact", "--omit=dev", "--no-audit", "--no-fund", spec], { env: installEnv, timeoutMs: 10 * 60_000 });
  if (result.code !== 0) throw new Error(result.errorCode === "ETIMEDOUT" ? "INSTALL_TIMEOUT" : "INSTALL_FAILED");
}
function entrypoint(prefix, item) {
  const root = path.join(prefix, "node_modules", item.package), pkg = store.readJson(path.join(root, "package.json"));
  if (pkg.version !== item.version) throw new Error("INSTALLED_VERSION_MISMATCH");
  const entry = typeof pkg.bin === "string" ? pkg.bin : pkg.bin?.[item.bin];
  if (typeof entry !== "string") throw new Error("CLI_ENTRYPOINT_MISSING");
  // OpenCode 1.x's Windows batch launcher uses `start`, which loses ACP's piped stdin.
  // Select the packaged native executable directly, retaining its exact platform package.
  if (item.bin === "opencode" && process.platform === "win32" && /\.cmd$/i.test(entry)) {
    const nativeName = `opencode-windows-${process.arch}`;
    let command;
    try { command = createRequire(path.join(root, "package.json")).resolve(`${nativeName}/bin/opencode.exe`); }
    catch { throw new Error("CLI_ENTRYPOINT_MISSING"); }
    const native = store.readJson(path.join(path.dirname(command), "..", "package.json"));
    if (native.name !== nativeName || native.version !== item.version) throw new Error("INSTALLED_VERSION_MISMATCH");
    if (!store.inside(prefix, fs.realpathSync(command))) throw new Error("CLI_ENTRYPOINT_INVALID");
    return command;
  }
  const command = path.resolve(root, entry);
  if (!store.inside(root, command) || !fs.existsSync(command) || !store.inside(prefix, fs.realpathSync(command))) throw new Error("CLI_ENTRYPOINT_INVALID");
  return command;
}
function fingerprint(file) { return store.fileDigest(file); }
function runtimeFiles(directory) {
  const files = {};
  function walk(current) {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      if (["profile.json", ".runtime-files.json"].includes(entry.name)) continue;
      const candidate = path.join(current, entry.name), relative = path.relative(directory, candidate);
      if (entry.isSymbolicLink()) {
        const target = fs.realpathSync(candidate);
        if (!store.inside(directory, target)) throw new Error("RUNTIME_EXTERNAL_LINK");
        files[relative] = { link: path.relative(directory, target) };
      } else if (entry.isDirectory()) walk(candidate);
      else files[relative] = { sha256: fingerprint(candidate) };
    }
  }
  walk(directory); return files;
}
export function verifyRuntime(profile, dataDir = store.home()) {
  const directory = path.join(dataDir, "runtimes", profile.id), files = store.readJson(path.join(directory, ".runtime-files.json"), 64 * 1024 * 1024);
  for (const [relative, expected] of Object.entries(files)) {
    const candidate = path.resolve(directory, relative);
    if (!store.inside(directory, candidate) || !fs.existsSync(candidate)) throw new Error("RUNTIME_INTEGRITY_FAILED");
    if (expected.link ? !store.inside(directory, fs.realpathSync(candidate)) || path.relative(directory, fs.realpathSync(candidate)) !== expected.link
      : fingerprint(candidate) !== expected.sha256) throw new Error("RUNTIME_INTEGRITY_FAILED");
  }
  return true;
}
async function verifyVersion(command, item, run, env) {
  const result = await run(command, ["--version"], { env: store.childEnvironment(env), timeoutMs: 15000 });
  const version = /\b(\d+\.\d+\.\d+(?:-[\w.-]+)?)\b/.exec(`${result.stdout}\n${result.stderr}`)?.[1];
  if (result.code !== 0 || version !== item.version) throw new Error("CLI_PREFLIGHT_FAILED");
}
async function nativePreflight(root, provider, command, { run, env }) {
  const result = await run(process.execPath, [path.join(root, "bin", "provider-preflight.mjs"), provider, command], { env: store.childEnvironment(env), timeoutMs: 90000 });
  let response;
  try { const line = result.stdout.split(/\r?\n/).find(line => line.startsWith("OFFICE_PARTNER_PREFLIGHT_RESULT:")); response = JSON.parse(line.slice("OFFICE_PARTNER_PREFLIGHT_RESULT:".length)); }
  catch { response = {}; }
  if (result.errorCode === "ETIMEDOUT") response.code = "TIMEOUT";
  if (result.code !== 0 || response.ok !== true) {
    const code = ["AUTH_REQUIRED", "POLICY_DENIED", "TIMEOUT", "PROTOCOL_INCOMPATIBLE", "INITIALIZATION_FAILED"].includes(response.code) ? response.code : "FAILED";
    throw new Error("NATIVE_PREFLIGHT_" + code);
  }
}
export async function captureExternal(root, provider, { dataDir = store.home(), env = process.env, run = runCommand } = {}) {
  const current = store.readState(dataDir);
  if (current.legacy[provider]) return;
  const command = resolveCommands(env[provider.toUpperCase() + "_BIN"] ?? provider, env)[0];
  if (!command) return;
  const result = await run(command, ["--version"], { env: store.childEnvironment(env), timeoutMs: 15000 });
  const version = /\b(\d+\.\d+\.\d+(?:-[\w.-]+)?)\b/.exec(`${result.stdout}\n${result.stderr}`)?.[1];
  if (result.code !== 0 || !version) return;
  const id = store.digest(JSON.stringify([provider, command, version, root]));
  store.changeState(dataDir, state => {
    state.profiles[id] = { id, provider, mode: "external", version, command, adapterRoot: root, fingerprint: fingerprint(command) };
    state.legacy[provider] = id;
  });
}
export async function installRuntime(root, provider, options = {}) {
  const { dataDir = store.home(), env = process.env, run = runCommand, install = installPackage, adapterRoot = root, preflight = nativePreflight } = options;
  if (!store.ids.includes(provider)) throw new Error("PROVIDER_INVALID");
  const item = readManifest(root).providers[provider];
  if (item.sdkVersion) {
    let entry;
    try { entry = createRequire(path.join(adapterRoot, "package.json")).resolve("@github/copilot-sdk"); } catch { throw new Error("SDK_DEPENDENCY_MISSING"); }
    let sdkDirectory = path.dirname(entry), sdkVersion;
    for (let i = 0; i < 8; i++, sdkDirectory = path.dirname(sdkDirectory)) {
      try { const pkg = store.readJson(path.join(sdkDirectory, "package.json")); if (pkg.name === "@github/copilot-sdk") { sdkVersion = pkg.version; break; } } catch { /* next */ }
    }
    if (sdkVersion !== item.sdkVersion) throw new Error("SDK_VERSION_MISMATCH");
  }
  const id = store.digest(JSON.stringify([provider, item, process.platform, process.arch, adapterRoot]));
  const directory = path.join(dataDir, "runtimes", id);
  const stateHome = path.join(dataDir, "provider-data", provider, `${item.version}-${process.platform}-${process.arch}`);
  const privateEnv = store.childEnvironment(env, { mode: "managed", provider, stateHome });
  return store.lock(dataDir, "install-" + provider, async () => {
    if (!store.inside(path.join(dataDir, "releases"), adapterRoot)) throw new Error("ADAPTER_SNAPSHOT_REQUIRED");
    await captureExternal(adapterRoot, provider, { dataDir, env, run });
    fs.mkdirSync(stateHome, { recursive: true, mode: 0o700 });
    let staging, fresh = false;
    try {
      if (!fs.existsSync(directory)) {
        fs.mkdirSync(path.dirname(directory), { recursive: true });
        staging = fs.mkdtempSync(path.join(path.dirname(directory), ".install-"));
        store.atomic(path.join(staging, "package.json"), { name: "office-partner-runtime", version: "1.0.0", private: true });
        await install(staging, `${item.package}@${item.version}`, { env: store.childEnvironment(env), run });
        await verifyVersion(entrypoint(staging, item), item, run, privateEnv);
        fs.renameSync(staging, directory); staging = null; fresh = true;
      }
      const command = entrypoint(directory, item);
      const profile = { id, provider, mode: "managed", version: item.version, sdkVersion: item.sdkVersion ?? null,
        command, adapterRoot, stateHome, fingerprint: fingerprint(command), platform: process.platform, arch: process.arch };
      if (!fs.existsSync(path.join(directory, ".runtime-files.json"))) {
        if (!fresh) throw new Error("RUNTIME_INTEGRITY_FAILED");
        store.atomic(path.join(directory, ".runtime-files.json"), runtimeFiles(directory));
      }
      verifyRuntime(profile, dataDir);
      await verifyVersion(command, item, run, privateEnv);
      store.atomic(path.join(directory, "profile.json"), profile);
      store.changeState(dataDir, state => {
        if (options.expectedPending && state.pending?.[provider] !== options.expectedPending) throw new Error("ACTIVATION_CANCELLED");
        state.profiles[id] = profile; if (options.trackPending !== false) { state.pending ??= {}; state.pending[provider] = id; }
      });
      await preflight(adapterRoot, provider, command, { run, env: privateEnv });
      if (options.canActivate && !options.canActivate()) throw new Error("ACTIVATION_CANCELLED");
      profile.verified = true;
      store.atomic(path.join(directory, "profile.json"), profile);
      store.changeState(dataDir, state => {
        if (options.expectedPending && state.pending?.[provider] !== options.expectedPending) throw new Error("ACTIVATION_CANCELLED");
        delete state.failures?.[provider];
        state.profiles[id] = profile;
        if (options.activate !== false) {
          if (state.active[provider] && state.active[provider] !== id) state.previous[provider] = state.active[provider];
          state.active[provider] = id;
          state.selected = [...new Set([...state.selected, provider])];
          state.disabled = (state.disabled ?? []).filter(value => value !== provider);
          delete state.pending?.[provider];
        }
      });
      return profile;
    } catch (error) {
      const known = /^(?:NATIVE_PREFLIGHT_[A-Z_]+|INSTALL_FAILED|INSTALL_TIMEOUT|CLI_PREFLIGHT_FAILED|INSTALLED_VERSION_MISMATCH|RUNTIME_INTEGRITY_FAILED|ACTIVATION_CANCELLED)$/;
      const code = known.test(error.message) ? error.message : "RUNTIME_INSTALL_FAILED";
      store.changeState(dataDir, state => { state.failures ??= {}; state.failures[provider] = { code, version: item.version, at: new Date().toISOString() }; });
      throw error;
    } finally {
      if (staging && store.inside(path.join(dataDir, "runtimes"), staging)) fs.rmSync(staging, { recursive: true, force: true });
    }
  });
}
export async function repairRuntime(root, provider, options = {}) {
  const dataDir = options.dataDir ?? store.home(), state = store.readState(dataDir), old = state.profiles[options.profileId ?? state.active[provider]];
  if (!old || old.mode !== "managed") throw new Error("NO_MANAGED_RUNTIME");
  // Never delete a directory that could be held by a worker. Build a fresh profile from a copied release.
  const repairedRoot = options.adapterRoot ?? old.adapterRoot;
  try { verifyRuntime(old, dataDir); return old; } catch { /* stage a complete replacement */ }
  // A damaged immutable installation is quarantined, not overwritten. Call only while services are stopped.
  const directory = path.join(dataDir, "runtimes", old.id);
  if (!store.inside(path.join(dataDir, "runtimes"), directory)) throw new Error("RUNTIME_PATH_INVALID");
  if (fs.existsSync(directory)) fs.renameSync(directory, directory + ".damaged-" + Date.now());
  return installRuntime(repairedRoot, provider, { ...options, dataDir, adapterRoot: repairedRoot,
    activate: old.id === state.active[provider], trackPending: old.id === state.active[provider] });
}
export function runtimeInventory(dataDir = store.home()) {
  const state = store.readState(dataDir);
  return store.ids.map(provider => {
    const profile = state.profiles[state.active[provider]];
    return { provider, selected: state.selected.includes(provider), disabled: state.disabled?.includes(provider) ?? false, mode: profile?.mode ?? "external", version: profile?.version ?? null,
      profileId: profile?.id ?? null, pendingVersion: state.profiles[state.pending?.[provider]]?.version ?? null, pendingVerified: state.profiles[state.pending?.[provider]]?.verified === true,
      previous: state.previous[provider] ?? null, available: profile ? fs.existsSync(profile.command) && fingerprint(profile.command) === profile.fingerprint : null };
  });
}
