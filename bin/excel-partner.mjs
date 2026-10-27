#!/usr/bin/env node
import { spawn, spawnSync } from "node:child_process";
import { X509Certificate } from "node:crypto";
import { copyFileSync, existsSync, mkdirSync, openSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import selfsigned from "selfsigned";
import { updatePackage } from "./update.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dataDir = path.join(process.env.LOCALAPPDATA || os.homedir(), "Excel Partner");
const cert = path.join(dataDir, "localhost.crt");
const key = path.join(dataDir, "localhost.key");
const pidFile = path.join(dataDir, "bridge.pid");
const configFile = path.join(dataDir, "config.json");
const manifest = path.join(root, "addin", "manifest.xml");
const bridge = path.join(root, "dist", "bridge", "src", "index.js");
const windowsScript = path.join(root, "scripts", "package-windows.ps1");

function fail(message) { throw new Error(message); }
function requireWindows() {
  if (process.platform !== "win32") fail("This package currently supports Excel desktop on Windows. Use the source WSL installer for WSL2.");
}
function powershell(action, extra = []) {
  const result = spawnSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", windowsScript,
    "-Action", action, "-DataDir", dataDir, "-ManifestPath", manifest,
    "-NodePath", process.execPath, "-CliPath", fileURLToPath(import.meta.url), ...extra], { stdio: "inherit" });
  if (result.error) throw result.error;
  if (result.status !== 0) fail(`Windows setup failed (${action}, exit ${result.status ?? "unknown"}).`);
}
function livePid() {
  if (!existsSync(pidFile)) return null;
  const pid = Number(readFileSync(pidFile, "utf8"));
  if (!Number.isSafeInteger(pid) || pid <= 0) return null;
  try { process.kill(pid, 0); return pid; } catch { return null; }
}
function bridgeEnv() {
  return { ...process.env, BRIDGE_CERT: cert, BRIDGE_KEY: key, BRIDGE_CONFIG: configFile,
    BRIDGE_USAGE_FILE: process.env.BRIDGE_USAGE_FILE ?? path.join(dataDir, "usage.jsonl") };
}
async function health() {
  return new Promise((resolve) => {
    if (!existsSync(cert)) return resolve(false);
    const req = https.get("https://127.0.0.1:8765/health", { ca: readFileSync(cert), timeout: 1000 }, (res) => {
      res.resume(); resolve(res.statusCode === 200);
    });
    req.on("error", () => resolve(false));
    req.on("timeout", () => { req.destroy(); resolve(false); });
  });
}
async function ensureCertificate() {
  if (existsSync(cert) && existsSync(key)) {
    const current = new X509Certificate(readFileSync(cert));
    if (new Date(current.validTo).getTime() > Date.now() + 30 * 86400000) return;
    fail(`Certificate expires soon. Run "excel-partner uninstall", then "excel-partner setup" to renew it.`);
  }
  if (existsSync(cert) || existsSync(key)) fail(`Incomplete certificate pair in ${dataDir}; remove it and rerun setup.`);
  mkdirSync(dataDir, { recursive: true });
  // Reuse the Agent Gateway's localhost certificate so the user trusts one certificate only.
  const shared = path.join(process.env.LOCALAPPDATA || os.homedir(), "Agent Gateway");
  const sharedCert = path.join(shared, "localhost.crt"), sharedKey = path.join(shared, "localhost.key");
  if (existsSync(sharedCert) && existsSync(sharedKey)
    && new Date(new X509Certificate(readFileSync(sharedCert)).validTo).getTime() > Date.now() + 30 * 86400000) {
    copyFileSync(sharedCert, cert); copyFileSync(sharedKey, key);
    console.log(`Reusing the Agent Gateway certificate from ${shared}`);
    return;
  }
  const pems = await selfsigned.generate([{ name: "commonName", value: "localhost" }], {
    algorithm: "sha256",
    notAfterDate: new Date(Date.now() + 365 * 86400000),
    extensions: [
      { name: "basicConstraints", cA: false },
      { name: "keyUsage", digitalSignature: true, keyEncipherment: true },
      { name: "extKeyUsage", serverAuth: true },
      { name: "subjectAltName", altNames: [{ type: 2, value: "localhost" }, { type: 7, ip: "127.0.0.1" }] },
    ],
  });
  writeFileSync(cert, pems.cert, { mode: 0o600 });
  writeFileSync(key, pems.private, { mode: 0o600 });
}
async function start() {
  if (!existsSync(cert) || !existsSync(key)) fail("Run 'excel-partner setup' first.");
  if (await health()) {
    if (!livePid()) fail("Port 8765 is already in use by another bridge. Stop it before starting this package.");
    console.log("Bridge is already running at https://localhost:8765"); return;
  }
  if (livePid()) fail(`Bridge process ${livePid()} exists but is not healthy. Run 'excel-partner stop' and retry.`);
  if (!existsSync(bridge)) fail("Package is missing its compiled bridge. Reinstall the package.");
  mkdirSync(dataDir, { recursive: true });
  const out = openSync(path.join(dataDir, "bridge.log"), "a");
  const err = openSync(path.join(dataDir, "bridge.err.log"), "a");
  const child = spawn(process.execPath, [bridge], { cwd: root, env: bridgeEnv(), detached: true, windowsHide: true,
    stdio: ["ignore", out, err] });
  child.unref();
  writeFileSync(pidFile, String(child.pid));
  for (let i = 0; i < 30; i++) {
    await new Promise((resolve) => setTimeout(resolve, 200));
    if (await health()) { console.log("Bridge ready at https://localhost:8765"); return; }
    if (!livePid()) break;
  }
  rmSync(pidFile, { force: true });
  fail(`Bridge did not start. See ${path.join(dataDir, "bridge.err.log")}`);
}
async function stop() {
  const pid = livePid();
  if (pid) {
    process.kill(pid);
    for (let i = 0; i < 25; i++) {
      try { process.kill(pid, 0); } catch { break; }
      await new Promise((resolve) => setTimeout(resolve, 200));
    }
    console.log(`Stopped bridge process ${pid}`);
  }
  rmSync(pidFile, { force: true });
}

const gatewayBin = path.join(root, "gateway", "bin", "agent-gateway.mjs");
const gatewayDir = process.env.AGENT_GATEWAY_HOME ?? path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".config"), "Agent Gateway");

/** Runs one Agent Gateway CLI command; its own output and exit code are passed through. */
function gateway(args) {
  if (!existsSync(gatewayBin)) fail("This installation is missing the gateway. Reinstall excel-partner.");
  const result = spawnSync(process.execPath, [gatewayBin, ...args], { stdio: "inherit" });
  if (result.error) throw result.error;
  if (result.status !== 0) { process.exitCode = result.status ?? 1; return false; }
  return true;
}
function tail(file, lines = 40) {
  if (!existsSync(file)) return [];
  return readFileSync(file, "utf8").split(/\r?\n/).filter(Boolean).slice(-lines);
}
function logsAddin() {
  for (const name of ["bridge.err.log", "bridge.log"]) {
    console.log(`--- ${path.join(dataDir, name)}`);
    console.log(tail(path.join(dataDir, name)).join("\n") || "(empty)");
  }
}
function readConfig() {
  try { return JSON.parse(readFileSync(configFile, "utf8")); } catch { return {}; }
}
/** Hostnames the bridge answers to besides localhost/127.0.0.1/::1; stored in config.json (env vars do not reach the sign-in startup entry). */
function allowHost(names, remove) {
  const config = readConfig();
  const current = Array.isArray(config.allowedHosts) ? config.allowedHosts.filter(name => typeof name === "string") : [];
  if (!names.length) {
    console.log(current.length ? `Extra allowed hostnames: ${current.join(", ")}` : "No extra hostnames allowed. The bridge answers to localhost, 127.0.0.1 and ::1.");
    return;
  }
  const clean = names.map(name => name.trim().toLowerCase().replace(/^\[(.*)\]$/, "$1"));
  for (const name of clean) if (!/^[a-z0-9.:_-]+$/.test(name) || name.length > 253) fail(`'${name}' is not a hostname or IP address (leave out http:// and the port).`);
  const next = remove ? current.filter(name => !clean.includes(name)) : [...new Set([...current, ...clean])];
  mkdirSync(dataDir, { recursive: true });
  writeFileSync(configFile, JSON.stringify({ ...config, allowedHosts: next }, null, 2) + "\n", { mode: 0o600 });
  console.log(`${remove ? "Removed" : "Allowed"}: ${clean.join(", ")}. Restart the bridge to apply: excel-partner restart`);
}
async function doctorAddin() {
  const line = (ok, text) => console.log(`${ok === true ? "[ ok ]" : ok === false ? "[fail]" : "[warn]"} ${text}`);
  line(Number(process.versions.node.split(".")[0]) >= 20, `Node.js ${process.versions.node} (20 or later required)`);
  const hasCert = existsSync(cert) && existsSync(key);
  line(hasCert, hasCert ? "Certificate present" : "Certificate missing. Run: excel-partner setup");
  if (hasCert) {
    const days = Math.floor((new Date(new X509Certificate(readFileSync(cert)).validTo).getTime() - Date.now()) / 86400000);
    line(days > 30 ? true : days > 0 ? null : false, `Certificate valid for ${days} more day(s)`);
  }
  const run = spawnSync("reg.exe", ["query", "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", "/v", "Excel Partner Bridge"], { stdio: "ignore", windowsHide: true });
  line(run.status === 0, run.status === 0 ? "Starts automatically at Windows sign-in" : "No sign-in startup entry. Run: excel-partner setup");
  const healthy = await health();
  line(healthy, healthy ? "Bridge running at https://localhost:8765" : "Bridge not running. Run: excel-partner start");
  const extra = readConfig().allowedHosts;
  line(true, Array.isArray(extra) && extra.length ? `Also answers to: ${extra.join(", ")}` : "Answers to localhost only. To reach it by another name (WSL IP): excel-partner allow-host <name>");
}

const TARGETS = { addin: "addin", app: "addin", gateway: "gateway", all: "all" };
const HELP = `Usage: excel-partner <command> [addin|gateway|all]

  setup      Register and start. No target = the Excel add-in. 'all' sets up both.
  start      Start the bridge (add-in) or the gateway.
  stop       Stop it.
  restart    Stop, then start.
  status     Show what is running (default: both).
  info       Show the gateway URL and token (add --copy to copy the token).
  logs       Show recent log lines.
  doctor     Check the installation and explain what to fix (default: both).
  allow-host Let the bridge answer to another hostname/IP (e.g. a WSL IP): allow-host <name> [--remove]; no name lists them.
  update     Install the latest release and refresh both (always the whole package).
  uninstall  Stop and remove registration and startup entries.

Gateway settings use: excel-partner gateway <config|alias|model-name|provider|models|warmup|autostart> ...
Examples: excel-partner setup            excel-partner setup gateway
          excel-partner start all        excel-partner info --copy`;

async function main() {
  const args = process.argv.slice(2).filter(arg => arg !== "--help");
  const command = args[0] || "help";
  if (command === "help" || process.argv.includes("--help")) { console.log(HELP); return; }
  if (command === "gateway") { gateway(args.slice(1)); return; }
  if (command === "allow-host") { allowHost(args.slice(1).filter(arg => !arg.startsWith("--")), args.includes("--remove")); return; }
  const flags = args.filter(arg => arg.startsWith("--"));
  const positional = args.filter(arg => !arg.startsWith("--"));
  const given = TARGETS[positional[1]];
  if (positional[1] && !given) fail(`Unknown target: ${positional[1]}. Use addin, gateway or all.`);
  const everything = ["status", "doctor", "update", "info", "token"].includes(command);
  const target = given ?? (everything ? "all" : "addin");
  const wantAddin = target !== "gateway", wantGateway = target !== "addin";
  // The add-in needs Excel desktop on Windows; the gateway does not. Each part
  // runs on its own, so a failure in one never prevents the other from running.
  const addin = fn => async () => { requireWindows(); await fn(); };
  const steps = [];
  const add = (name, fn) => steps.push([name, fn]);
  const gatewayStep = args => () => { if (!gateway(args)) throw new Error("see the output above"); };

  switch (command) {
    case "setup":
      if (wantAddin) add("Add-in", addin(async () => {
        await ensureCertificate();
        powershell("register");
        await stop();
        await start();
        console.log("Restart Excel, then open Home > Excel Partner. Use Settings > Sign in to GitHub Copilot for device-code sign-in; other providers use their own CLI login.");
      }));
      if (wantGateway) add("Gateway", gatewayStep(["setup"]));
      break;
    case "start":
      if (wantAddin) add("Add-in", addin(start));
      if (wantGateway) add("Gateway", gatewayStep(["start"]));
      break;
    case "stop":
      if (wantGateway) add("Gateway", gatewayStep(["stop"]));
      if (wantAddin) add("Add-in", addin(stop));
      break;
    case "restart":
      if (wantGateway) add("Gateway", gatewayStep(["restart"]));
      if (wantAddin) add("Add-in", addin(async () => { await stop(); await start(); }));
      break;
    case "status":
      if (wantAddin) add("Add-in", async () => console.log((await health()) ? "Add-in bridge: running at https://localhost:8765" : "Add-in bridge: stopped or unhealthy"));
      if (wantGateway) add("Gateway", gatewayStep(["status"]));
      break;
    case "info":
    case "token":
      if (wantAddin && target === "all") add("Add-in", async () => console.log("Add-in bridge: https://localhost:8765 (the add-in connects automatically; no token needed)"));
      if (wantGateway) add("Gateway", () => {
        if (!existsSync(path.join(gatewayDir, "token.txt"))) console.log("Gateway: not set up. Run: excel-partner setup gateway");
        else gatewayStep([command, ...flags])();
      });
      break;
    case "logs":
      if (wantAddin) add("Add-in", logsAddin);
      if (wantGateway) add("Gateway", gatewayStep(["logs"]));
      break;
    case "doctor":
      if (wantAddin) add("Add-in", async () => { console.log("Excel add-in"); await doctorAddin(); });
      if (wantGateway) add("Gateway", () => { console.log("\nAgent Gateway"); gatewayStep(["doctor"])(); });
      break;
    case "update":
      add("Update", addin(async () => {
        // The gateway ships in the same package: stop it while files are replaced, then start the new one.
        const gatewaySetUp = existsSync(path.join(gatewayDir, "token.txt"));
        if (gatewaySetUp) gateway(["stop"]);
        try { await updatePackage({ root, dataDir, nodePath: process.execPath, stop, wasRunning: Boolean(livePid()) }); }
        finally {
          if (gatewaySetUp) spawnSync(process.execPath, [fileURLToPath(import.meta.url), "gateway", "start"], { stdio: "inherit" });
        }
      }));
      break;
    case "uninstall":
      if (wantGateway) add("Gateway", gatewayStep(["uninstall"]));
      if (wantAddin) add("Add-in", addin(async () => {
        await stop();
        powershell("unregister");
        rmSync(cert, { force: true });
        rmSync(key, { force: true });
        console.log(`Registration removed. User data remains at ${dataDir}`);
      }));
      break;
    default:
      fail(`Unknown command: ${command}. Run excel-partner help.`);
  }
  for (const [name, step] of steps) {
    try { await step(); }
    catch (error) { console.error(`${name}: ${error.message || error}`); process.exitCode = 1; }
  }
}
main().catch((error) => { console.error(error.message || error); process.exitCode = 1; });
