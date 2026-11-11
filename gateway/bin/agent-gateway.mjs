#!/usr/bin/env node
import { spawn, spawnSync } from "node:child_process";
import { randomBytes, X509Certificate } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import https from "node:https";
import { fileURLToPath, pathToFileURL } from "node:url";
import { ConfigSchema, dataDir, loadConfig } from "../dist/config.mjs";
import { decideRecorded, processOwner } from "./process-record.mjs";

const dir = dataDir();
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const cert = path.join(dir, "localhost.crt"), key = path.join(dir, "localhost.key");
const tokenFile = path.join(dir, "token.txt"), configFile = path.join(dir, "config.json"), pidFile = path.join(dir, "server.json");
const entry = path.join(root, "dist", "server.mjs");
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
async function waitForGateway(ms) {
  for (let waited = 0; waited < ms; waited += 200) {
    if (await health()) return true;
    await delay(200);
  }
  return false;
}
function token() { return fs.readFileSync(tokenFile, "utf8").trim(); }
function request(method, route, auth = false) {
  return new Promise((resolve, reject) => {
    const req = https.request({ hostname: "localhost", port: loadConfig(dir).port, path: route, method, ca: fs.readFileSync(cert),
      headers: auth ? { "x-api-key": token() } : {} }, res => {
      let text = ""; res.setEncoding("utf8"); res.on("data", chunk => text += chunk); res.on("end", () => {
        try { const body = JSON.parse(text); if (res.statusCode !== 200) reject(new Error(body.error?.message ?? `HTTP ${res.statusCode}`)); else resolve(body); }
        catch (err) { reject(err); }
      });
    });
    req.on("error", reject); req.setTimeout(route === "/v1/models" || route === "/gateway/warmup" ? 180_000 : 2000, () => req.destroy(new Error("Gateway did not respond"))); req.end();
  });
}
async function health() { try { return await request("GET", "/health"); } catch { return null; } }
function save(config) {
  const checked = ConfigSchema.parse(config);
  if (!checked.enabledProviders.includes(checked.defaultRoute.provider)) throw new Error("Default provider must remain enabled");
  fs.mkdirSync(dir, { recursive: true });
  const temporary = `${configFile}.tmp`;
  fs.writeFileSync(temporary, JSON.stringify(checked, null, 2), { mode: 0o600 }); fs.renameSync(temporary, configFile);
  console.log("Configuration saved.");
}
/** Apply a saved configuration to a running gateway without a separate restart command. */
async function reload() {
  if (!await health()) return;
  console.log("Restarting the gateway to apply it…");
  await stop(); await start();
}
function copy(text) {
  return process.platform === "win32" && spawnSync("clip.exe", { input: text, windowsHide: true }).status === 0;
}
function tail(file, lines = 40) {
  if (!fs.existsSync(file)) return [];
  return fs.readFileSync(file, "utf8").split(/\r?\n/).filter(Boolean).slice(-lines);
}
function logs() {
  for (const name of ["gateway.err.log", "gateway.log"]) {
    console.log(`--- ${path.join(dir, name)}`);
    console.log(tail(path.join(dir, name)).join("\n") || "(empty)");
  }
}
async function doctor() {
  const line = (ok, text) => console.log(`${ok === true ? "[ ok ]" : ok === false ? "[fail]" : "[warn]"} ${text}`);
  line(Number(process.versions.node.split(".")[0]) >= 20, `Node.js ${process.versions.node} (20 or later required)`);
  let config;
  try { config = loadConfig(dir); line(true, `Configuration valid (port ${config.port})`); }
  catch (error) { line(false, `Configuration invalid: ${error.message.split("\n")[0]}. Fix ${configFile}`); return; }
  const hasFiles = fs.existsSync(cert) && fs.existsSync(key) && fs.existsSync(tokenFile);
  line(hasFiles, hasFiles ? "Certificate and token present" : "Not set up. Run: office-partner setup gateway");
  if (!hasFiles) return;
  const x509 = new X509Certificate(fs.readFileSync(cert));
  const days = Math.floor((new Date(x509.validTo).getTime() - Date.now()) / 86400000);
  line(days > 30 ? true : days > 0 ? null : false, `Certificate valid for ${days} more day(s)`);
  if (process.platform === "win32") {
    const store = spawnSync("certutil.exe", ["-user", "-store", "Root"], { encoding: "utf8", windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
    const trusted = (store.stdout ?? "").toLowerCase().includes(x509.fingerprint.replaceAll(":", "").toLowerCase());
    line(trusted, trusted ? "Certificate trusted for this Windows user" : "Certificate not trusted. Run: office-partner setup gateway");
    const run = spawnSync("reg.exe", ["query", runKey, "/v", runName], { stdio: "ignore", windowsHide: true });
    line(run.status === 0 ? true : null, run.status === 0 ? "Starts automatically at Windows sign-in" : "No sign-in startup entry. Run: office-partner gateway autostart on");
  }
  const h = await health();
  line(h?.ready === true, h ? `Gateway running at https://localhost:${config.port}` : "Gateway not running. Run: office-partner start gateway");
  if (!h) return;
  try {
    const status = await request("GET", "/gateway/status", true);
    for (const p of status.providers) line(p.status === "ready" ? true : p.installed ? null : false, `${p.provider}: ${p.status}${p.hint ? ` - ${p.hint}` : ""}`);
  } catch (error) { line(false, `Authenticated request failed: ${error.message}. Check the token with: office-partner info`); }
}
async function uninstall() {
  await stop();
  autostart(false);
  console.log(`Gateway stopped and its startup entry removed. Configuration, token and certificate remain in ${dir}.\nTo remove the program: npm uninstall -g office-partner`);
}
// Reuse the Excel Partner add-in's localhost certificate so the user trusts one certificate only.
function adoptCertificate(from) {
  const c = path.join(from, "localhost.crt"), k = path.join(from, "localhost.key");
  if (!fs.existsSync(c) || !fs.existsSync(k)) return;
  if (new Date(new X509Certificate(fs.readFileSync(c)).validTo).getTime() < Date.now() + 30 * 86400000) return;
  fs.copyFileSync(c, cert); fs.copyFileSync(k, key);
  console.log(`Reusing the Excel Partner certificate from ${from}`);
}
async function setup() {
  fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(configFile)) save({});
  if (!fs.existsSync(tokenFile)) fs.writeFileSync(tokenFile, `agw-${randomBytes(32).toString("hex")}`, { mode: 0o600 });
  if (fs.existsSync(cert) !== fs.existsSync(key)) throw new Error("Incomplete TLS certificate pair; restore both files before setup");
  if (!fs.existsSync(cert)) adoptCertificate(path.join(process.env.LOCALAPPDATA ?? os.homedir(), "Excel Partner"));
  if (!fs.existsSync(cert)) {
    const { default: selfsigned } = await import("selfsigned");
    const pems = await selfsigned.generate([{ name: "commonName", value: "localhost" }], { algorithm: "sha256", notAfterDate: new Date(Date.now() + 365 * 86400000), extensions: [
      { name: "basicConstraints", cA: false }, { name: "keyUsage", digitalSignature: true, keyEncipherment: true },
      { name: "extKeyUsage", serverAuth: true }, { name: "subjectAltName", altNames: [{ type: 2, value: "localhost" }, { type: 7, ip: "127.0.0.1" }, { type: 7, ip: "::1" }] },
    ] });
    fs.writeFileSync(cert, pems.cert, { mode: 0o600 }); fs.writeFileSync(key, pems.private, { mode: 0o600 });
  }
  if (new Date(new X509Certificate(fs.readFileSync(cert)).validTo).getTime() < Date.now()) throw new Error("Local TLS certificate expired");
  if (process.platform === "win32") {
    const result = spawnSync("certutil.exe", ["-addstore", "-user", "Root", cert], { stdio: "inherit", windowsHide: true });
    if (result.status !== 0) throw new Error("Could not trust the local certificate");
  } else console.log(`Trust this certificate in your browser/OS: ${cert}`);
  await start();
  autostart(true);
  info(true);
  // One real minimal inference now, so the first Office login is not limited by a cold CLI.
  try { await request("POST", "/gateway/warmup", true); console.log("Model verification cache warmed."); }
  catch (error) { console.log(`Warmup skipped (${error.message}). Run "agent-gateway warmup" before signing in from Office.`); }
}
const runKey = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", runName = "Agent Gateway";
function autostart(enable) {
  if (process.platform !== "win32") { console.log("Automatic startup is only managed on Windows."); return; }
  const startup = process.env.OFFICE_PARTNER_RECOVERY_ENTRY
    ? `"${process.execPath}" "${process.env.OFFICE_PARTNER_RECOVERY_ENTRY}" start gateway --at-logon`
    : `"${process.execPath}" "${fileURLToPath(import.meta.url)}" start`;
  const args = enable ? ["add", runKey, "/v", runName, "/t", "REG_SZ", "/d", startup, "/f"]
    : ["delete", runKey, "/v", runName, "/f"];
  const result = spawnSync("reg.exe", args, { stdio: "ignore", windowsHide: true });
  if (result.status !== 0 && enable) throw new Error("Could not register Windows sign-in startup");
  console.log(enable ? "Gateway will start automatically when you sign in to Windows." : "Automatic startup removed.");
}
function info(copyToken = false) {
  console.log(`Gateway URL: https://localhost:${loadConfig(dir).port}\nToken: ${token()}\nConfiguration: ${configFile}`);
  if (copyToken) console.log(copy(token()) ? "Token copied to the clipboard." : "Could not copy the token; select it above.");
}
async function start(repair = false) {
  if (!fs.existsSync(cert) || !fs.existsSync(key) || !fs.existsSync(tokenFile)) throw new Error("Run agent-gateway setup first");
  const existing = await health();
  if (existing) {
    if (existing.service !== "agent-gateway") throw new Error("Port is occupied by another service");
    console.log(`Gateway already running at https://localhost:${loadConfig(dir).port}`); return;
  }
  if (fs.existsSync(pidFile)) {
    let previous = null; try { previous = JSON.parse(fs.readFileSync(pidFile, "utf8")); } catch { /* A damaged record is just stale. */ }
    let alive = false; try { if (previous) { process.kill(previous.pid, 0); alive = true; } } catch {}
    if (alive) {
      // A process number alone proves nothing: after a restart Windows gives the same number to another program.
      const owner = processOwner(previous.pid, line => typeof previous.entry === "string" && previous.entry ? line.toLowerCase().includes(previous.entry.toLowerCase()) : /server\.mjs/i.test(line));
      const action = await decideRecorded(owner, { repair, waitHealthy: () => waitForGateway(30_000) });
      if (action === "ready") { console.log(`Gateway already running at https://localhost:${loadConfig(dir).port}`); return; }
      if (action === "fail") throw new Error('Recorded gateway is still alive but unhealthy. Open Start > "Office Partner - Fix connection", or inspect its logs before restarting');
      if (action === "kill") {
        process.kill(previous.pid);
        for (let i = 0; i < 25; i++) { try { process.kill(previous.pid, 0); } catch { break; } await delay(200); }
      }
    }
    fs.rmSync(pidFile, { force: true });
  }
  const instance = randomBytes(24).toString("hex");
  const out = fs.openSync(path.join(dir, "gateway.log"), "a"), err = fs.openSync(path.join(dir, "gateway.err.log"), "a");
  let child;
  try { child = spawn(process.execPath, [entry], { cwd: dir, env: { ...process.env, AGENT_GATEWAY_HOME: dir, AGENT_GATEWAY_INSTANCE: instance },
    detached: true, windowsHide: true, stdio: ["ignore", out, err] }); }
  finally { fs.closeSync(out); fs.closeSync(err); }
  child.on("error", e => console.error(e.message)); child.unref();
  fs.writeFileSync(pidFile, JSON.stringify({ pid: child.pid, instance, entry }), { mode: 0o600 });
  for (let i = 0; i < Math.ceil(loadConfig(dir).requestTimeoutMs / 200) + 50; i++) {
    await delay(200); const h = await health();
    if (h?.instance === instance && h.ready) { console.log(`Gateway ready at https://localhost:${loadConfig(dir).port}`); return; }
    if (child.exitCode !== null) break;
  }
  throw new Error(`Gateway failed to start. Inspect ${path.join(dir, "gateway.err.log")}`);
}
async function stop() {
  const h = await health();
  if (!h) { console.log("Gateway is stopped or unhealthy; no process was killed."); return; }
  if (!fs.existsSync(pidFile)) throw new Error("No matching process record; refusing to stop an unmanaged service");
  const record = JSON.parse(fs.readFileSync(pidFile, "utf8"));
  if (record.instance !== h.instance || h.service !== "agent-gateway") throw new Error("Process identity mismatch; refusing to stop a different service");
  await request("POST", "/gateway/shutdown", true);
  for (let i = 0; i < 50; i++) {
    await delay(200);
    let alive = false; try { process.kill(record.pid, 0); alive = true; } catch {}
    if (!await health() && !alive) { fs.rmSync(pidFile, { force: true }); console.log("Gateway stopped."); return; }
  }
  throw new Error("Gateway is still shutting down; inspect logs before retrying");
}
async function main() {
  const argv = process.argv.slice(2);
  const copyToken = argv.includes("--copy");
  const repair = argv.includes("--repair");
  const [command = "status", a, b, c, d] = argv.filter(arg => arg !== "--copy" && arg !== "--repair");
  if (process.env.OFFICE_PARTNER_GATEWAY_MANAGED !== "1") {
    const packageRoot = path.dirname(root);
    if (["setup", "uninstall"].includes(command)) {
      const result = spawnSync(process.execPath, [path.join(packageRoot, "bin", "office-partner.mjs"), command, "gateway"], { stdio: "inherit", windowsHide: true });
      process.exitCode = result.status ?? 1; return;
    }
    const { dispatchActive } = await import(pathToFileURL(path.join(packageRoot, "bin", "manage.mjs")).href);
    if (dispatchActive(packageRoot, ["gateway", ...argv])) return;
  }
  if (command === "setup") await setup();
  else if (command === "start") await start(repair);
  else if (command === "stop") await stop();
  else if (command === "restart") { await stop(); await start(repair); }
  else if (command === "status") {
    if (!fs.existsSync(tokenFile)) { console.log("Gateway is not set up. Run: office-partner setup gateway"); return; }
    const h = await health(); console.log(h ? `Gateway running at https://localhost:${loadConfig(dir).port} (${h.sessions} active sessions)` : "Gateway stopped or unhealthy. Run: office-partner doctor gateway");
  }
  else if (command === "info") info(copyToken);
  else if (command === "token") { console.log(token()); if (copyToken) console.error(copy(token()) ? "Copied to the clipboard." : "Could not copy."); }
  else if (command === "logs") logs();
  else if (command === "doctor") await doctor();
  else if (command === "uninstall") await uninstall();
  else if (command === "update") {
    const result = spawnSync(process.execPath, [path.join(root, "..", "bin", "office-partner.mjs"), "update"], { stdio: "inherit" });
    process.exitCode = result.status ?? 1;
  }
  else if (command === "autostart" && ["on", "off"].includes(a)) autostart(a === "on");
  else if (command === "models") { const result = await request("GET", "/v1/models", true); for (const model of result.data) console.log(`${model.id}\t${model.display_name}`); }
  else if (command === "warmup") { await request("POST", "/gateway/warmup", true); console.log("Model verification cache warmed using a real CLI inference."); }
  else if (command === "config" && a === "show") console.log(JSON.stringify(loadConfig(dir), null, 2));
  else if (command === "config" && a === "set" && b && c !== undefined) {
    const config = loadConfig(dir); let value; try { value = JSON.parse(c); } catch { value = c; }
    const keys = b.split("."); if (keys.some(k => ["__proto__", "prototype", "constructor"].includes(k))) throw new Error("Invalid configuration key");
    let object = config;
    for (const key of keys.slice(0, -1)) { if (!object[key] || typeof object[key] !== "object") throw new Error("Unknown configuration key"); object = object[key]; }
    if (!(keys.at(-1) in object)) throw new Error("Unknown configuration key");
    object[keys.at(-1)] = value; save(config); await reload();
  } else if (command === "alias" && a === "set" && b && c && d) {
    if (["__proto__", "prototype", "constructor"].includes(b)) throw new Error("Invalid alias");
    const config = loadConfig(dir); config.aliases[b] = { provider: c, model: d === "default" ? null : d, effort: config.defaultRoute.effort }; save(config); await reload();
  } else if (command === "model-name" && a === "set" && b && c) {
    if (["__proto__", "prototype", "constructor"].includes(b)) throw new Error("Invalid model ID");
    const config = loadConfig(dir); config.modelNames[b] = c; save(config); await reload();
  } else if (command === "provider" && ["enable", "disable"].includes(a) && b) {
    const config = loadConfig(dir); config.enabledProviders = a === "enable" ? [...new Set([...config.enabledProviders, b])] : config.enabledProviders.filter(p => p !== b); save(config); await reload();
  } else {
    if (command !== "help" && command !== "--help") throw new Error("Unknown or incomplete command. Run agent-gateway help.");
    console.log("agent-gateway setup|start|stop|restart|status|info [--copy]|token [--copy]|logs|doctor|update|uninstall\nagent-gateway models|warmup\nagent-gateway autostart on|off\nSettings apply immediately to a running gateway.\nagent-gateway config show\nagent-gateway config set <key> <value>\nagent-gateway alias set <api-model> <provider> <native-model|default>\nagent-gateway model-name set <model-id> <friendly-name>\nagent-gateway provider enable|disable <provider>");
  }
}
main().catch(e => { console.error(e.message); process.exitCode = 1; });
