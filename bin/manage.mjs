import fs from "node:fs";
import path from "node:path";
import https from "node:https";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import { createRequire } from "node:module";
import store from "./runtime-store.cjs";
import { collectDoctor, resolveCommands } from "./doctor-checks.mjs";
import { installRuntime, repairRuntime, runtimeInventory } from "./runtime-manager.mjs";
import { snapshotRelease, updateRelease, rollbackRelease, verifyRelease, recoverInterrupted, restoreSelection } from "./deployment.mjs";

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
        for (let i = 0; i < 50; i++) { await new Promise(resolve => setTimeout(resolve, 100)); if ((await control(dataDir, role, "health")).absent) { gone = true; break; } }
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
  if (state.owners.addin || state.owners.gateway || !gateway && gatewayConfigured) return;
  const deployment = store.currentDeployment(dataDir);
  deployment.active = null; deployment.previous = null; deployment.journal = null;
  store.atomic(path.join(dataDir, "deployment.json"), deployment);
  for (const profile of Object.values(state.profiles)) {
    if (profile.mode !== "managed") continue;
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
  store.atomic(path.join(dataDir, "runtime-state.json"), { schemaVersion: 1, selected: [], profiles: {}, active: {}, previous: {}, legacy: {}, conversations: {}, owners: {} });
  store.atomic(path.join(dataDir, "deployment.json"), { schemaVersion: 1, active: null, previous: null, releases: {}, journal: null });
}
export function verifiedActiveRelease(dataDir = store.home()) {
  const deployment = store.currentDeployment(dataDir), release = deployment.releases[deployment.active];
  if (!release) return null;
  try { verifyRelease(release); return release; }
  catch {
    const previous = deployment.releases[deployment.previous];
    if (!previous || previous.id === release.id) throw new Error("ACTIVE_RELEASE_DAMAGED_USE_RECOVERY");
    verifyRelease(previous, true);
    restoreSelection(dataDir, previous.runtimeSelection ?? (previous.runtimeProfiles ? { active: previous.runtimeProfiles } : undefined));
    deployment.active = previous.id; deployment.previous = null; deployment.held = true;
    deployment.releases[release.id].failed = true;
    deployment.journal = { phase: "integrity_fallback", from: release.id, to: previous.id, updated: new Date().toISOString() };
    store.atomic(path.join(dataDir, "deployment.json"), deployment);
    return previous;
  }
}
export function dispatchActive(root, args, dataDir = store.home()) {
  if (process.env.OFFICE_PARTNER_DISPATCHED === "1") return false;
  recoverInterrupted(dataDir);
  const release = verifiedActiveRelease(dataDir);
  if (!release || path.resolve(root) === path.resolve(release.root)) return false;
  const result = spawnSync(process.execPath, [path.join(release.root, "bin", "office-partner.mjs"), ...args], {
    env: { ...process.env, OFFICE_PARTNER_DISPATCHED: "1", OFFICE_PARTNER_DATA_DIR: dataDir, OFFICE_PARTNER_RECOVERY_ENTRY: path.join(dataDir, "recovery", "recovery.mjs") },
    stdio: "inherit", windowsHide: true,
  });
  process.exitCode = result.status ?? 1; return true;
}
export async function managementMain(args, options = {}) {
  const dataDir = options.dataDir ?? store.home();
  const [command = "versions", action, value] = args;
  if (["start", "restart"].includes(command)) { recoverInterrupted(dataDir); verifiedActiveRelease(dataDir); }
  if (command === "repair") {
    for (const [name, validate] of [["runtime-state.json", store.readState], ["deployment.json", store.currentDeployment]]) {
      try { validate(dataDir); }
      catch {
        if (!fs.existsSync(path.join(dataDir, name + ".previous"))) throw new Error("RECOVERY_METADATA_BACKUP_MISSING");
        const previous = validate(dataDir, ".previous"); store.atomic(path.join(dataDir, name), previous);
      }
    }
    recoverInterrupted(dataDir);
  }
  const deployment = store.currentDeployment(dataDir);
  const active = deployment.releases[deployment.active];
  const root = active?.root ?? options.root;
  if (command === "versions") {
    console.log(JSON.stringify({ active: active?.version ?? null, previous: deployment.releases[deployment.previous]?.version ?? null, held: deployment.held ?? false,
      cached: Object.values(deployment.releases).map(release => ({ id: release.id, version: release.version, created: release.created })), providers: runtimeInventory(dataDir) }, null, 2)); return;
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
      const baseline = ensureBaseline(root, dataDir);
      const profile = await installRuntime(root, value, { dataDir, adapterRoot: baseline.root });
      console.log(`Installed managed ${value} ${profile.version}. Existing workers retain their runtime.`); return;
    }
    if (action === "remove") {
      store.changeState(dataDir, state => { state.selected = state.selected.filter(id => id !== value); delete state.active[value]; delete state.pending?.[value]; state.disabled = [...new Set([...(state.disabled ?? []), value])]; });
      console.log("Provider disabled. Retained profiles remain available to saved conversations until full uninstall."); return;
    }
    if (action === "login") {
      const state = store.readState(dataDir);
      const profile = state.profiles[state.pending?.[value]] ?? store.selectProfile(dataDir, value)
        ?? { mode: "external", command: resolveCommands(process.env[value.toUpperCase() + "_BIN"] ?? value, process.env)[0], provider: value };
      if (!profile.command) throw new Error("INSTALL_PROVIDER_FIRST");
      const login = value === "copilot" ? ["login", "--device-code"] : value === "codex" ? ["login"] : value === "claude" ? ["auth", "login"] : ["auth", "login"];
      const js = /\.[cm]?js$/i.test(profile.command);
      let interactive = spawnSync;
      try { interactive = createRequire(path.join(root, "package.json"))("cross-spawn").sync; } catch { /* native/JS entrypoints still work for recovery */ }
      const result = interactive(js ? process.execPath : profile.command, js ? [profile.command, ...login] : login, { env: store.childEnvironment(process.env, profile), stdio: "inherit", windowsHide: true });
      process.exitCode = result.status ?? 1;
      if (result.status === 0 && state.pending?.[value]) await installRuntime(profile.adapterRoot, value, { dataDir, adapterRoot: profile.adapterRoot, activate: active?.root === profile.adapterRoot, expectedPending: profile.id });
      return;
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
    const details = { ...report, deployment: { active: active?.version ?? null, previous: deployment.releases[deployment.previous]?.version ?? null,
      phase: deployment.journal?.phase ?? null, held: deployment.held ?? false }, providers: runtimeInventory(dataDir) };
    fs.writeFileSync(value, JSON.stringify(details, null, 2) + "\n", { mode: 0o600, flag: "wx" });
    console.log("Saved diagnostic metadata. No raw logs, credentials or document contents were included."); return;
  }
  if (["start", "stop", "restart", "status", "doctor", "uninstall"].includes(command)) {
    const result = spawnSync(process.execPath, [path.join(root, "bin", "office-partner.mjs"), ...args], { env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir, OFFICE_PARTNER_DISPATCHED: "1" }, stdio: "inherit", windowsHide: true });
    process.exitCode = result.status ?? 1; return;
  }
  const controls = options.lifecycle ?? await lifecycle(dataDir);
  const report = options.report ?? (phase => console.log(`Deployment: ${phase}`));
  if (command === "rollback") { const restored = await rollbackRelease(action, { dataDir, ...controls, report }); console.log("Previous complete release restored."); return restored; }
  if (command === "repair") {
    if (!active) throw new Error("INSTALLATION_NOT_INITIALIZED");
    try { verifyRelease(active, true); }
    catch {
      if (deployment.previous) {
        const restored = await rollbackRelease(undefined, { dataDir, ...controls, report, allowDamagedCurrent: true });
        console.log("Damaged release replaced by the retained complete release."); return restored;
      }
      throw new Error("RELEASE_DAMAGED_NO_VERIFIED_BACKUP_REINSTALL_REQUIRED");
    }
    return store.lock(dataDir, "deployment", async () => {
      let stopped = false, started = false;
      try {
        await controls.drain(); stopped = true;
        await controls.stop();
        // Saved native conversations and cached rollbacks can still need inactive profiles.
        for (const profile of Object.values(store.readState(dataDir).profiles)) {
          if (profile.mode === "managed") await repairRuntime(profile.adapterRoot, profile.provider, { dataDir, profileId: profile.id });
        }
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
    if (!spec || spec === "office-partner@latest") {
      const response = await fetch("https://registry.npmjs.org/office-partner/latest", { signal: AbortSignal.timeout(15000) });
      if (!response.ok) throw new Error("RELEASE_LOOKUP_FAILED");
      const latest = (await response.json()).version;
      if (!/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(latest)) throw new Error("RELEASE_VERSION_INVALID");
      spec = "office-partner@" + latest;
    }
    if (/\.tgz$/i.test(spec)) spec = path.resolve(spec);
    installRecovery(root, dataDir);
    const updated = await updateRelease(root, spec, { dataDir, ...controls, report }); console.log("Release and selected provider runtimes updated together."); return updated;
  }
  throw new Error("MANAGEMENT_COMMAND_INVALID");
}
