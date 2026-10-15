#!/usr/bin/env node
import { spawn, spawnSync } from "node:child_process";
import { X509Certificate } from "node:crypto";
import { existsSync, mkdirSync, openSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import selfsigned from "selfsigned";

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
  return { ...process.env, BRIDGE_CERT: cert, BRIDGE_KEY: key, BRIDGE_CONFIG: configFile };
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

async function main() {
  const command = process.argv[2] || "help";
  if (command === "help" || command === "--help") {
    console.log("Usage: excel-partner <setup|start|stop|status|uninstall>\nExcel desktop on Windows; setup registers the local add-in for the current user.");
    return;
  }
  requireWindows();
  if (command === "setup") {
    await ensureCertificate();
    powershell("register");
    await stop();
    await start();
    console.log("Restart Excel, then open Home > Excel Partner. Use Settings > Sign in to GitHub Copilot for device-code sign-in; other providers use their own CLI login.");
  } else if (command === "start") await start();
  else if (command === "stop") await stop();
  else if (command === "status") console.log((await health()) ? "Bridge running at https://localhost:8765" : "Bridge stopped or unhealthy");
  else if (command === "uninstall") {
    await stop();
    powershell("unregister");
    rmSync(cert, { force: true });
    rmSync(key, { force: true });
    console.log(`Registration removed. User data remains at ${dataDir}`);
  } else fail(`Unknown command: ${command}`);
}
main().catch((error) => { console.error(error.message || error); process.exitCode = 1; });
