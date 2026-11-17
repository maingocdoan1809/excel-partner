import fs from "node:fs";
import path from "node:path";
import https from "node:https";
import { spawnSync } from "node:child_process";
import store from "./runtime-store.cjs";
import { collectDoctor } from "./doctor-checks.mjs";
import { installRuntime, runtimeInventory, uninstallRuntime } from "./runtime-manager.mjs";
import { recordStartup, startupCode } from "./startup-log.mjs";

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
const START_ATTEMPTS = 3;
const START_PAUSE_MS = 15_000;
/**
 * One try at starting a service. Never throws: the outcome is returned so it can be logged,
 * because the sign-in entry runs without a window and nobody would see an error.
 */
function tryStart(args, dataDir, options) {
  try {
    const root = options.root;
    if (!root) return { ok: false, message: "INSTALLATION_NOT_INITIALIZED" };
    const launch = options.launch ?? (() => spawnSync(process.execPath, [path.join(root, "bin", "office-partner.mjs"), ...args], {
      env: { ...process.env, OFFICE_PARTNER_DATA_DIR: dataDir },
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
  for (const [service, label, args] of steps) {
    const outcome = tryStart(args, dataDir, options);
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
  const [command, action, value] = args;
  if (command === "start") return startServices(args, dataDir, options);
  const root = options.root;
  if (!root) throw new Error("INSTALLATION_NOT_INITIALIZED");
  if (command === "fix") return fixConnection(dataDir, options);
  if (command === "providers") {
    if (!action || action === "list") { console.log(JSON.stringify(runtimeInventory(dataDir), null, 2)); return; }
    if (!store.ids.includes(value)) throw new Error("PROVIDER_INVALID");
    if (action === "install") {
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
    const details = { ...report, version: store.readJson(path.join(root, "package.json")).version, providers: runtimeInventory(dataDir) };
    fs.writeFileSync(value, JSON.stringify(details, null, 2) + "\n", { mode: 0o600, flag: "wx" });
    console.log("Saved diagnostic metadata. No raw logs, credentials or document contents were included."); return;
  }
  throw new Error("MANAGEMENT_COMMAND_INVALID");
}
