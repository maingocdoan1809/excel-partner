import fs from "node:fs";
import path from "node:path";
import https from "node:https";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import store from "./runtime-store.cjs";
import { collectDoctor } from "./doctor-checks.mjs";
import { installRuntime, repairRuntime, runtimeInventory, uninstallRuntime } from "./runtime-manager.mjs";
import { recordStartup, startupCode } from "./startup-log.mjs";
import { snapshotRelease, updateRelease, verifyRelease } from "./deployment.mjs";

export function controlRequest(dataDir, role, action, timeoutMs = 5000) {
  const directory = role === "addin" ? dataDir : process.env.AGENT_GATEWAY_HOME ?? path.join(process.env.LOCALAPPDATA || path.dirname(dataDir), "Agent Gateway");
  const config = role === "gateway" && fs.existsSync(path.join(directory, "config.json")) ? store.readJson(path.join(directory, "config.json")) : {};
  const port = role === "addin" ? Number(process.env.BRIDGE_PORT ?? 8765) : config.port ?? 8766;
  return new Promise(resolve => {
    const cert = path.join(directory, "localhost.crt");
    if (!fs.existsSync(cert)) return resolve({ absent: true });
    let token;
    try { token = fs.readFileSync(path.join(directory, role === "addin" ? "control-token.txt" : "token.txt"), "utf8").trim(); } catch { /* absent/legacy */ }
    const route = action === "health" ? "/health" : `${role === "addin" ? "/admin" : "/gateway"}/${action}`;
    let settled = false, body = "", timer;
    const done = value => { if (!settled) { settled = true; clearTimeout(timer); resolve(value); } };
    const request = https.request({ hostname: "127.0.0.1", port, path: route, method: action === "health" ? "GET" : "POST", ca: fs.readFileSync(cert),
      headers: token ? { [role === "addin" ? "x-office-partner-control" : "x-api-key"]: token } : {} }, response => {
      response.on("data", chunk => { body += chunk; if (body.length > 256 * 1024) { done({ error: "RESPONSE_LIMIT" }); request.destroy(); } });
      response.on("error", () => done({ error: "SERVICE_RESPONSE_FAILED" }));
      response.on("end", () => {
        try {
          const data = JSON.parse(body);
          if (action === "health") {
            const valid = role === "addin" ? data.name === "excel-ai-addin-bridge" : data.service === "agent-gateway";
            done(valid ? { running: true, controlled: role === "gateway" || data.controlSupported === true,
              busy: role === "addin" ? data.busy === true || Array.isArray(data.sessions) && data.sessions.some(session => session.activeRequestId || session.pendingTools > 0) : (data.busySessions ?? data.sessions ?? 0) > 0 } : { error: "SERVICE_IDENTITY_MISMATCH" });
          } else done(response.statusCode === 200 ? { ok: true } : { error: response.statusCode === 409 ? "SERVICE_BUSY" : "SERVICE_CONTROL_UNAVAILABLE" });
        } catch { done({ error: "SERVICE_RESPONSE_INVALID" }); }
      });
    });
    request.on("error", error => done(error.code === "ECONNREFUSED" ? { absent: true } : { error: "SERVICE_CONNECTION_FAILED" }));
    timer = setTimeout(() => { done({ error: "SERVICE_CONTROL_TIMEOUT" }); request.destroy(); }, timeoutMs);
    request.end();
  });
}
export async function lifecycle(dataDir, { control = controlRequest, run = spawnSync } = {}) {
  const roles = [];
  const owners = store.readState(dataDir).owners;
  for (const role of ["addin", "gateway"]) {
    const status = await control(dataDir, role, "health");
    if (status.error) throw new Error(status.error);
    if (status.running || owners[role]) roles.push(role);
  }
  return {
    drain: async () => {
      for (const role of roles) { const status = await control(dataDir, role, "health"); if (status.absent) continue;
        const result = await control(dataDir, role, "maintenance"); if (!result.ok) throw new Error(result.error ?? "SERVICE_DRAIN_FAILED"); }
    },
    undrain: async () => { for (const role of roles) await control(dataDir, role, "resume"); },
    stop: async () => {
      for (const role of [...roles].reverse()) {
        const status = await control(dataDir, role, "health");
        if (status.absent) continue;
        if (status.error) throw new Error(status.error);
        const stopped = await control(dataDir, role, "shutdown");
        if (!stopped.ok) throw new Error(stopped.error ?? "SERVICE_STOP_FAILED");
        let gone = false;
        for (let i = 0; i < 1800; i++) { await new Promise(resolve => setTimeout(resolve, 100)); if ((await control(dataDir, role, "health")).absent) { gone = true; break; } }
        if (!gone) throw new Error("SERVICE_STOP_TIMEOUT");
      }
    },
    start: async release => {
      for (const role of roles) {
        const result = run(process.execPath, [path.join(release.root, "bin", "office-partner.mjs"), "setup", role, "--skip-providers"], {
          cwd: dataDir, env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir, OFFICE_PARTNER_DISPATCHED: "1", OFFICE_PARTNER_RECOVERY_ENTRY: path.join(dataDir, "recovery", "recovery.mjs") },
          windowsHide: true, stdio: "inherit", timeout: 180000,
        });
        if (result.error || result.status !== 0) throw new Error("SERVICE_START_FAILED");
        const checked = await control(dataDir, role, "health"); if (!checked.running) throw new Error("SERVICE_HEALTH_FAILED");
      }
    },
  };
}
export function installRecovery(root, dataDir = store.home()) {
  const recovery = path.join(dataDir, "recovery"); fs.mkdirSync(recovery, { recursive: true });
  const source = path.join(root, "bin"), files = fs.readdirSync(source).filter(file => fs.statSync(path.join(source, file)).isFile()).sort();
  const id = store.digest(files.map(file => file + ":" + store.fileDigest(path.join(source, file))).join("\n"));
  const destination = path.join(recovery, "versions", id);
  if (!fs.existsSync(destination)) {
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    const staging = fs.mkdtempSync(path.join(recovery, ".install-"));
    try { fs.cpSync(source, path.join(staging, "bin"), { recursive: true }); fs.renameSync(staging, destination); }
    finally { if (fs.existsSync(staging)) fs.rmSync(staging, { recursive: true, force: true }); }
  }
  store.atomic(path.join(recovery, "package.json"), { name: "office-partner-recovery", type: "module" });
  const launcher = "import path from 'node:path'; import {fileURLToPath} from 'node:url'; import { managementMain } from './versions/" + id + "/bin/manage.mjs';\nmanagementMain(process.argv.slice(2),{dataDir:path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..')}).catch(error=>{ console.error(error.code || error.message || 'RECOVERY_FAILED'); process.exitCode=1; });\n";
  const temporary = path.join(recovery, crypto.randomUUID() + ".tmp");
  try { fs.writeFileSync(temporary, launcher, { mode: 0o600 }); fs.renameSync(temporary, path.join(recovery, "recovery.mjs")); }
  finally { if (fs.existsSync(temporary)) fs.unlinkSync(temporary); }
  if (!fs.existsSync(path.join(dataDir, "control-token.txt"))) fs.writeFileSync(path.join(dataDir, "control-token.txt"), crypto.randomBytes(32).toString("hex"), { mode: 0o600, flag: "wx" });
}
export function ensureBaseline(root, dataDir = store.home()) {
  let deployment = store.currentDeployment(dataDir);
  if (deployment.active) return deployment.releases[deployment.active];
  const baseline = snapshotRelease(root, dataDir); deployment = store.currentDeployment(dataDir); deployment.active = baseline.id;
  store.atomic(path.join(dataDir, "deployment.json"), deployment); return baseline;
}
export function releaseOwners(dataDir, options) {
  return store.lock(dataDir, "deployment", () => {
    const locks = (remaining, work) => remaining.length ? store.lock(dataDir, "install-" + remaining[0], () => locks(remaining.slice(1), work)) : work();
    return locks(store.ids, () => cleanupOwners(dataDir, options));
  });
}
function cleanupOwners(dataDir, { addin, gateway, gatewayConfigured = false }) {
  store.changeState(dataDir, state => { if (addin) state.owners.addin = false; if (gateway) state.owners.gateway = false; });
  const state = store.readState(dataDir);
  if (state.owners.addin || state.owners.gateway || !gateway && gatewayConfigured && state.owners.gateway !== false) return;
  const deployment = store.currentDeployment(dataDir);
  deployment.active = null;
  store.atomic(path.join(dataDir, "deployment.json"), deployment);
  for (const profile of Object.values(state.profiles)) {
    const runtimeRoot = path.join(dataDir, "runtimes");
    for (const name of fs.existsSync(runtimeRoot) ? fs.readdirSync(runtimeRoot) : []) {
      if (name !== profile.id && !new RegExp("^" + profile.id + "\\.damaged-\\d+$").test(name)) continue;
      const directory = path.join(runtimeRoot, name);
      if (!store.inside(runtimeRoot, fs.realpathSync(directory)) || store.readJson(path.join(directory, "profile.json")).id !== profile.id) throw new Error("UNINSTALL_OWNERSHIP_MISMATCH");
      fs.rmSync(directory, { recursive: true, force: true });
    }
  }
  for (const release of Object.values(deployment.releases)) {
    if (!fs.existsSync(release.directory)) continue;
    if (!store.inside(path.join(dataDir, "releases"), fs.realpathSync(release.directory)) || store.readJson(path.join(release.directory, ".release.json"), 64 * 1024 * 1024).id !== release.id) throw new Error("UNINSTALL_OWNERSHIP_MISMATCH");
    fs.rmSync(release.directory, { recursive: true, force: true });
  }
  store.atomic(path.join(dataDir, "runtime-state.json"), store.emptyState());
  store.atomic(path.join(dataDir, "deployment.json"), { schemaVersion: 1, active: null, releases: {} });
}
export function verifiedActiveRelease(dataDir = store.home()) {
  const deployment = store.currentDeployment(dataDir), release = deployment.releases[deployment.active];
  if (!release) return null;
  // No older release is kept to fall back to. A damaged one is replaced by installing it again (`repair`).
  try { verifyRelease(release); return release; }
  catch { throw new Error("ACTIVE_RELEASE_DAMAGED_USE_RECOVERY"); }
}
export function dispatchActive(root, args, dataDir = store.home()) {
  if (process.env.OFFICE_PARTNER_DISPATCHED === "1") return false;
  const release = verifiedActiveRelease(dataDir);
  if (!release || path.resolve(root) === path.resolve(release.root)) return false;
  const result = spawnSync(process.execPath, [path.join(release.root, "bin", "office-partner.mjs"), ...args], {
    env: { ...process.env, OFFICE_PARTNER_DISPATCHED: "1", OFFICE_PARTNER_DATA_DIR: dataDir, OFFICE_PARTNER_RECOVERY_ENTRY: path.join(dataDir, "recovery", "recovery.mjs") },
    stdio: "inherit", windowsHide: true,
  });
  process.exitCode = result.status ?? 1; return true;
}
const START_ATTEMPTS = 3;
const START_PAUSE_MS = 15_000;
/**
 * One try at starting a service through the active release. Never throws: the outcome is returned so it can be logged,
 * because the sign-in entry runs without a window and nobody would see an error.
 */
function tryStart(args, dataDir, options) {
  try {
    verifiedActiveRelease(dataDir);
    const deployment = store.currentDeployment(dataDir);
    const root = deployment.releases[deployment.active]?.root ?? options.root;
    if (!root) return { ok: false, message: "INSTALLATION_NOT_INITIALIZED" };
    const launch = options.launch ?? (() => spawnSync(process.execPath, [path.join(root, "bin", "office-partner.mjs"), ...args], {
      env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir, OFFICE_PARTNER_DISPATCHED: "1" },
      stdio: ["inherit", "inherit", "pipe"], encoding: "utf8", windowsHide: true }));
    const result = launch(root, args);
    if (result.stderr) process.stderr.write(result.stderr);
    if (result.error) return { ok: false, message: String(result.error.message) };
    return { ok: result.status === 0, message: result.status === 0 ? "" : (result.stderr ?? "").trim() || `exit ${result.status}` };
  } catch (error) { return { ok: false, message: String(error?.code ?? error?.message ?? error) }; }
}
/** `start [addin|gateway|all]`. At sign-in (--at-logon) a failed start is tried again; every attempt is recorded. */
async function startServices(args, dataDir, options) {
  const atLogon = args.includes("--at-logon");
  const forwarded = args.filter(arg => arg !== "--at-logon");
  const target = forwarded.find((arg, index) => index > 0 && !arg.startsWith("--")) ?? "addin";
  const services = target === "all" ? ["addin", "gateway"] : [target];
  for (let attempt = 1; attempt <= (atLogon ? START_ATTEMPTS : 1); attempt++) {
    const outcome = tryStart(forwarded, dataDir, options);
    for (const service of services) recordStartup(dataDir, { service, attempt, ok: outcome.ok, ...(outcome.ok ? {} : { code: startupCode(outcome.message), message: outcome.message }) });
    if (outcome.ok) return;
    if (atLogon && attempt < START_ATTEMPTS) await new Promise(resolve => setTimeout(resolve, options.startPauseMs ?? START_PAUSE_MS));
  }
  process.exitCode = 1;
}
/** The Start Menu shortcut "Office Partner - Fix connection": bring the bridge and, if set up, the gateway back, and say what happened. */
async function fixConnection(dataDir, options) {
  const gatewayDir = process.env.AGENT_GATEWAY_HOME ?? path.join(process.env.LOCALAPPDATA || path.dirname(dataDir), "Agent Gateway");
  const steps = [["addin", "Office Partner (Word, Excel, PowerPoint)", ["start", "--repair"]]];
  if (fs.existsSync(path.join(gatewayDir, "token.txt"))) steps.push(["gateway", "Gateway for the Claude Office add-ins", ["start", "gateway", "--repair"]]);
  console.log("Checking Office Partner…");
  let failed = 0;
  let reinstalled = false;
  for (const [service, label, args] of steps) {
    let outcome = tryStart(args, dataDir, options);
    if (!outcome.ok && !reinstalled && startupCode(outcome.message) === "RELEASE_DAMAGED") {
      // No older release is kept to fall back to, so the damaged one is installed again from npm.
      reinstalled = true;
      console.log("The installed files are damaged. Installing the same version again from npm…");
      try { await managementMain(["repair"], options); outcome = tryStart(args, dataDir, options); }
      catch (error) { outcome = { ok: false, message: String(error?.code ?? error?.message ?? error) }; }
    }
    recordStartup(dataDir, { service, attempt: 1, ok: outcome.ok, source: "fix", ...(outcome.ok ? {} : { code: startupCode(outcome.message), message: outcome.message }) });
    console.log(`${outcome.ok ? "OK    " : "FAILED"} ${label}${outcome.ok ? "" : `: ${startupCode(outcome.message)}`}`);
    if (!outcome.ok) failed++;
  }
  console.log(failed
    ? `\nSomething is still not working. Send this file to support:\n${path.join(dataDir, "startup.log")}`
    : "\nDone. Close the Office Partner pane and open it again (Home > Office Partner). If it still shows an error, restart Word, Excel or PowerPoint.");
  process.exitCode = failed ? 1 : 0;
  // The shortcut opens its own window; keep it open until the person has read the result.
  if (process.stdin.isTTY) {
    const { createInterface } = await import("node:readline/promises");
    const input = createInterface({ input: process.stdin, output: process.stdout });
    try { await input.question("\nPress Enter to close this window."); } finally { input.close(); }
  }
}
const PRODUCT = "office-partner";
const VERSION = /^\d+\.\d+\.\d+(?:-[\w.-]+)?$/;

/** `npm view <spec> <field> --json`: npm's own registry, proxy, certificate and auth settings apply (Node's fetch ignores them). */
async function npmView(spec, field) {
  const { npmEntrypoint } = await import("./runtime-manager.mjs");
  const { runCommand } = await import("./doctor-probes.mjs");
  const result = await runCommand(process.execPath, [npmEntrypoint(), "view", spec, field, "--json"], { timeoutMs: 30_000 });
  if (result.code !== 0) throw new Error("RELEASE_LOOKUP_FAILED");
  const text = result.stdout.trim();
  return text ? JSON.parse(text) : undefined;
}

export function compareVersions(a, b) {
  const parse = version => { const [core, pre] = version.split("-", 2); return { parts: core.split(".").map(Number), pre }; };
  const x = parse(a), y = parse(b);
  for (let i = 0; i < 3; i++) if (x.parts[i] !== y.parts[i]) return x.parts[i] > y.parts[i] ? 1 : -1;
  if (x.pre === y.pre) return 0;
  return x.pre === undefined ? 1 : y.pre === undefined ? -1 : x.pre > y.pre ? 1 : -1;
}

/**
 * What `update` installs, under the admin policy: updates can be turned off or pinned to one version, and the registry
 * chosen. Otherwise "latest" from npm; never an older version (unless asked, as rollback does) or a deprecated one.
 */
export async function updateTarget(root, spec, { allowDowngrade = false, view = npmView, policy = store.readPolicy() } = {}) {
  if (!policy.autoUpdate) throw new Error(policy.error ? "UPDATE_DISABLED_POLICY_INVALID" : "UPDATE_DISABLED_BY_POLICY");
  if (policy.registry) process.env.npm_config_registry = policy.registry; // npm view, pack and install all read it
  const latest = !spec || spec === PRODUCT + "@latest";
  if (policy.pinVersion) {
    if (!latest && spec !== PRODUCT + "@" + policy.pinVersion) throw new Error("UPDATE_PINNED");
    return PRODUCT + "@" + policy.pinVersion; // the administrator's choice wins, including going back
  }
  if (/\.tgz$/i.test(spec ?? "")) return spec; // a file the user picked on purpose
  if (latest) spec = PRODUCT + "@" + await view(PRODUCT + "@latest", "version");
  const target = spec.startsWith(PRODUCT + "@") ? spec.slice(PRODUCT.length + 1) : null;
  if (!target || !VERSION.test(target)) throw new Error("RELEASE_VERSION_INVALID");
  const current = store.readJson(path.join(root, "package.json")).version;
  if (VERSION.test(current ?? "") && compareVersions(target, current) < 0 && !allowDowngrade) throw new Error("UPDATE_DOWNGRADE_REFUSED");
  // A version picked by hand is the user's choice; "latest" that the publisher deprecated (a withdrawn release) is not taken.
  if (latest && await view(spec, "deprecated")) throw new Error("UPDATE_VERSION_DEPRECATED");
  return spec;
}

export async function managementMain(args, options = {}) {
  const dataDir = options.dataDir ?? store.home();
  const [command = "versions", action, value] = args;
  if (command === "start") return startServices(args, dataDir, options);
  if (command === "fix") return fixConnection(dataDir, options);
  if (command === "restart") verifiedActiveRelease(dataDir);
  if (command === "repair") {
    for (const [name, validate] of [["runtime-state.json", store.readState], ["deployment.json", store.currentDeployment]]) {
      try { validate(dataDir); }
      catch {
        if (!fs.existsSync(path.join(dataDir, name + ".previous"))) throw new Error("RECOVERY_METADATA_BACKUP_MISSING");
        const previous = validate(dataDir, ".previous"); store.atomic(path.join(dataDir, name), previous);
      }
    }
  }
  const deployment = store.currentDeployment(dataDir);
  const active = deployment.releases[deployment.active];
  const root = active?.root ?? options.root;
  if (command === "versions") {
    console.log(JSON.stringify({ active: active?.version ?? null, providers: runtimeInventory(dataDir) }, null, 2)); return;
  }
  if (!root) throw new Error("INSTALLATION_NOT_INITIALIZED");
  if (command === "doctor") {
    const { doctorMain } = await import("./doctor.mjs");
    process.exitCode = await doctorMain(args.slice(1), { root, env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir } }); return;
  }
  if (command === "providers") {
    if (!action || action === "list") { console.log(JSON.stringify(runtimeInventory(dataDir), null, 2)); return; }
    if (!store.ids.includes(value)) throw new Error("PROVIDER_INVALID");
    if (action === "install") {
      ensureBaseline(root, dataDir);
      const profile = await installRuntime(root, value, { dataDir });
      console.log(`Installed managed ${value} ${profile.version}.`); return;
    }
    if (action === "remove") {
      uninstallRuntime(value, { dataDir });
      console.log("Provider disabled."); return;
    }
    throw new Error("PROVIDER_COMMAND_INVALID");
  }
  if (command === "debug") {
    if (!["on", "off"].includes(action)) throw new Error("USE_DEBUG_ON_OR_OFF");
    store.atomic(path.join(dataDir, "debug.json"), { expires: action === "on" ? Date.now() + 15 * 60000 : 0 });
    console.log(action === "on" ? "Debug logging enabled for 15 minutes." : "Debug logging disabled."); return;
  }
  if (command === "diagnostics") {
    if (action !== "export" || !value) throw new Error("USE_DIAGNOSTICS_EXPORT_NEW_FILE_JSON");
    const report = await collectDoctor({ root, env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir } });
    const details = { ...report, deployment: { active: active?.version ?? null }, providers: runtimeInventory(dataDir) };
    fs.writeFileSync(value, JSON.stringify(details, null, 2) + "\n", { mode: 0o600, flag: "wx" });
    console.log("Saved diagnostic metadata. No raw logs, credentials or document contents were included."); return;
  }
  if (["stop", "restart", "status", "doctor", "uninstall"].includes(command)) {
    const result = spawnSync(process.execPath, [path.join(root, "bin", "office-partner.mjs"), ...args], { env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir, OFFICE_PARTNER_DISPATCHED: "1" }, stdio: "inherit", windowsHide: true });
    process.exitCode = result.status ?? 1; return;
  }
  if (command === "rollback") {
    // Older releases are not kept on disk: going back means installing that version from npm.
    if (!/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(action ?? "")) throw new Error("USE_ROLLBACK_WITH_A_VERSION");
    return managementMain(["update", "--from", "office-partner@" + action, "--allow-downgrade"], options);
  }
  const controls = options.lifecycle ?? await lifecycle(dataDir);
  const startedAt = Date.now();
  const report = options.report ?? (phase => console.log(`Deployment: ${phase} (${Math.round((Date.now() - startedAt) / 1000)}s)`));
  if (command === "repair") {
    if (!active) throw new Error("INSTALLATION_NOT_INITIALIZED");
    try { verifyRelease(active, true); }
    catch {
      // No older release is kept, so a damaged one is replaced by installing the same version again from npm.
      const reinstalled = await updateRelease(root, "office-partner@" + active.version, { dataDir, ...controls, report, ...options.deploy });
      console.log("Damaged release replaced by a fresh install of the same version."); return reinstalled;
    }
    return store.lock(dataDir, "deployment", async () => {
      let stopped = false, started = false;
      try {
        await controls.drain(); stopped = true;
        await controls.stop();
        for (const provider of Object.keys(store.readState(dataDir).active)) await repairRuntime(active.root, provider, { dataDir });
        await controls.start(active); started = true;
      } finally {
        try { if (stopped && !started) await controls.start(active); }
        finally { await controls.undrain(); }
      }
      console.log("Release integrity checked and managed runtimes repaired."); return active;
    });
  }
  if (command === "update") {
    let spec = action === "--from" ? value : undefined;
    spec = await updateTarget(root, spec, { allowDowngrade: args.includes("--allow-downgrade"), view: options.npmView });
    if (/\.tgz$/i.test(spec)) spec = path.resolve(spec);
    installRecovery(root, dataDir);
    const updated = await updateRelease(root, spec, { dataDir, ...controls, report, notice: message => console.log(message), ...options.deploy }); console.log("Release and selected provider runtimes updated together."); return updated;
  }
  throw new Error("MANAGEMENT_COMMAND_INVALID");
}
