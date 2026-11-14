import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import store from "./runtime-store.cjs";
import { resolveCommands } from "./doctor-checks.mjs";
import { runCommand } from "./doctor-probes.mjs";

const VERSION = /^\d+\.\d+\.\d+(?:-[\w.-]+)?$/;
export function readManifest(root) {
  const manifest = store.readJson(path.join(root, "bin", "runtime-manifest.json"));
  if (manifest.schemaVersion !== 1 || !manifest.providers) throw new Error("RUNTIME_MANIFEST_INVALID");
  for (const id of store.ids) {
    const item = manifest.providers[id];
    if (!item || typeof item.package !== "string" || item.bin !== id || !VERSION.test(item.version) || item.sdkVersion && !/^\d+\.\d+\.\d+$/.test(item.sdkVersion)) throw new Error("RUNTIME_MANIFEST_INVALID");
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
  const result = await run(process.execPath, [npmEntrypoint(env), "install", "--global=false", "--prefix", prefix, "--save-exact", "--omit=dev", "--no-audit", "--no-fund", "--prefer-offline", spec], { env: installEnv, timeoutMs: 10 * 60_000 });
  if (result.code !== 0) throw new Error(result.errorCode === "ETIMEDOUT" ? "INSTALL_TIMEOUT" : "INSTALL_FAILED");
}
/**
 * The file that starts the CLI. A package may ship a native executable for the platform instead of a script launcher
 * (`nativeBin` in the manifest): the launcher can lose the piped stdin that ACP needs, so the executable is used directly.
 */
function entrypoint(prefix, item) {
  const root = path.join(prefix, "node_modules", item.package), pkg = store.readJson(path.join(root, "package.json"));
  if (pkg.version !== item.version) throw new Error("INSTALLED_VERSION_MISMATCH");
  const entry = typeof pkg.bin === "string" ? pkg.bin : pkg.bin?.[item.bin];
  if (typeof entry !== "string") throw new Error("CLI_ENTRYPOINT_MISSING");
  const native = item.nativeBin?.[process.platform]?.replace("{arch}", process.arch);
  if (native && /\.cmd$/i.test(entry)) {
    const [nativePackage] = native.split("/");
    let command;
    try { command = createRequire(path.join(root, "package.json")).resolve(native); }
    catch { throw new Error("CLI_ENTRYPOINT_MISSING"); }
    const installed = store.readJson(path.join(command, ...Array(native.split("/").length - 1).fill(".."), "package.json"));
    if (installed.name !== nativePackage || installed.version !== item.version) throw new Error("INSTALLED_VERSION_MISMATCH");
    if (!store.inside(prefix, fs.realpathSync(command))) throw new Error("CLI_ENTRYPOINT_INVALID");
    return command;
  }
  const command = path.resolve(root, entry);
  if (!store.inside(root, command) || !fs.existsSync(command) || !store.inside(prefix, fs.realpathSync(command))) throw new Error("CLI_ENTRYPOINT_INVALID");
  return command;
}
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
      else files[relative] = { sha256: store.fileDigest(candidate) };
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
      : store.fileDigest(candidate) !== expected.sha256) throw new Error("RUNTIME_INTEGRITY_FAILED");
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

/** Where the CLI keeps its own sign-in, settings and chats. It is kept from one CLI version to the next, so an update never signs anyone out. */
function stateHomeFor(dataDir, provider, state) {
  return state.profiles[state.active[provider]]?.stateHome ?? path.join(dataDir, "provider-data", provider, "home");
}

/**
 * Installs the CLI the manifest of `root` names for `provider`, checks it and makes it the one the provider runs (`activate: false` only
 * records it). The install folder is named by what is installed, so a release that ships the same CLI reuses it instead of downloading it again.
 */
export async function installRuntime(root, provider, options = {}) {
  const { dataDir = store.home(), env = process.env, run = runCommand, install = installPackage, preflight = nativePreflight, activate = true } = options;
  if (!store.ids.includes(provider)) throw new Error("PROVIDER_INVALID");
  const item = readManifest(root).providers[provider];
  if (item.sdkVersion) {
    let entry;
    try { entry = createRequire(path.join(root, "package.json")).resolve("@github/copilot-sdk"); } catch { throw new Error("SDK_DEPENDENCY_MISSING"); }
    let sdkDirectory = path.dirname(entry), sdkVersion;
    for (let i = 0; i < 8; i++, sdkDirectory = path.dirname(sdkDirectory)) {
      try { const pkg = store.readJson(path.join(sdkDirectory, "package.json")); if (pkg.name === "@github/copilot-sdk") { sdkVersion = pkg.version; break; } } catch { /* next */ }
    }
    if (sdkVersion !== item.sdkVersion) throw new Error("SDK_VERSION_MISMATCH");
  }
  const state = store.readState(dataDir);
  // The CLI the provider already runs is kept when the release names the same package and version: nothing is downloaded or checked
  // again, whatever else in the manifest changed (a new check, a new variable). Only a different CLI gets a new install folder.
  const current = state.profiles[state.active[provider]];
  const sameCli = current && current.provider === provider && current.version === item.version && (current.sdkVersion ?? null) === (item.sdkVersion ?? null)
    && current.platform === process.platform && current.arch === process.arch && fs.existsSync(path.join(dataDir, "runtimes", current.id, "node_modules", item.package, "package.json"));
  // `minVersion` only changes what is checked, so it is not part of what is installed.
  const { minVersion: _checkedOnly, ...installed } = item;
  const id = sameCli ? current.id : store.digest(JSON.stringify([provider, installed, process.platform, process.arch]));
  const directory = path.join(dataDir, "runtimes", id);
  const stateHome = stateHomeFor(dataDir, provider, state);
  const privateEnv = store.childEnvironment(env, { provider, stateHome });
  return store.lock(dataDir, "install-" + provider, async () => {
    fs.mkdirSync(stateHome, { recursive: true, mode: 0o700 });
    let staging, fresh = false;
    try {
      if (!fs.existsSync(directory)) {
        fs.mkdirSync(path.dirname(directory), { recursive: true });
        staging = fs.mkdtempSync(path.join(path.dirname(directory), ".install-"));
        store.atomic(path.join(staging, "package.json"), { name: "office-partner-runtime", version: "1.0.0", private: true });
        await install(staging, `${item.package}@${item.version}`, { env: store.childEnvironment(env), run });
        await verifyVersion(entrypoint(staging, item), item, run, privateEnv);
        store.renameFolder(staging, directory); staging = null; fresh = true;
      }
      const command = entrypoint(directory, item);
      // The CLI the provider already runs was checked when it was installed (files, version, a session with the add-in's tools), and
      // `doctor` and `repair` check it again. An update that only changes the add-in's own code does not repeat that: it can take half a minute.
      const profile = sameCli ? current : { id, provider, version: item.version, sdkVersion: item.sdkVersion ?? null, command, stateHome,
        fingerprint: store.fileDigest(command), platform: process.platform, arch: process.arch };
      if (!sameCli) {
        if (fresh) store.atomic(path.join(directory, ".runtime-files.json"), runtimeFiles(directory));
        else if (!fs.existsSync(path.join(directory, ".runtime-files.json"))) throw new Error("RUNTIME_INTEGRITY_FAILED");
        else { verifyRuntime(profile, dataDir); await verifyVersion(command, item, run, privateEnv); }
        store.atomic(path.join(directory, "profile.json"), profile);
        // Whether a session opens with the add-in's tools needs the person to be signed in. Not being signed in yet is not a reason to refuse the CLI.
        try { await preflight(root, provider, command, { run, env: privateEnv }); }
        catch (error) { if (error.message !== "NATIVE_PREFLIGHT_AUTH_REQUIRED") throw error; }
      }
      store.changeState(dataDir, state => {
        state.profiles[id] = profile;
        delete state.failures[provider];
        if (!activate) return;
        state.active[provider] = id;
        state.selected = [...new Set([...state.selected, provider])];
        state.disabled = state.disabled.filter(value => value !== provider);
        dropUnused(state);
      });
      return profile;
    } catch (error) {
      const known = /^(?:NATIVE_PREFLIGHT_[A-Z_]+|INSTALL_FAILED|INSTALL_TIMEOUT|CLI_PREFLIGHT_FAILED|INSTALLED_VERSION_MISMATCH|RUNTIME_INTEGRITY_FAILED)$/;
      const code = known.test(error.message) ? error.message : "RUNTIME_INSTALL_FAILED";
      store.changeState(dataDir, state => { state.failures[provider] = { code, version: item.version, at: new Date().toISOString() }; });
      throw error;
    } finally {
      if (staging && store.inside(path.join(dataDir, "runtimes"), staging)) fs.rmSync(staging, { recursive: true, force: true });
    }
  });
}
/** Stops using the CLI the product installed for `provider`, and keeps it from being used until it is installed again. */
export function uninstallRuntime(provider, { dataDir = store.home() } = {}) {
  if (!store.ids.includes(provider)) throw new Error("PROVIDER_INVALID");
  store.changeState(dataDir, state => {
    state.selected = state.selected.filter(id => id !== provider);
    delete state.active[provider];
    state.disabled = [...new Set([...state.disabled, provider])];
    dropUnused(state);
  });
}
/** Forgets the installs no provider runs. Their folders are removed once nothing can be using them (after an update has settled). */
export function dropUnused(state) {
  const active = new Set(Object.values(state.active));
  for (const id of Object.keys(state.profiles)) if (!active.has(id)) delete state.profiles[id];
}
export async function repairRuntime(root, provider, options = {}) {
  const dataDir = options.dataDir ?? store.home(), state = store.readState(dataDir), old = state.profiles[state.active[provider]];
  if (!old) throw new Error("NO_MANAGED_RUNTIME");
  try { verifyRuntime(old, dataDir); return old; } catch { /* install a complete replacement */ }
  // A damaged immutable installation is quarantined, not overwritten. Call only while services are stopped.
  const directory = path.join(dataDir, "runtimes", old.id);
  if (!store.inside(path.join(dataDir, "runtimes"), directory)) throw new Error("RUNTIME_PATH_INVALID");
  if (fs.existsSync(directory)) store.renameFolder(directory, directory + ".damaged-" + Date.now());
  return installRuntime(root, provider, { ...options, dataDir });
}
export function runtimeInventory(dataDir = store.home()) {
  const state = store.readState(dataDir);
  return store.ids.map(provider => {
    const profile = state.profiles[state.active[provider]];
    return { provider, selected: state.selected.includes(provider), disabled: state.disabled.includes(provider), managed: Boolean(profile), version: profile?.version ?? null,
      profileId: profile?.id ?? null, failure: state.failures[provider]?.code ?? null,
      available: profile ? fs.existsSync(profile.command) && store.fileDigest(profile.command) === profile.fingerprint : null };
  });
}
