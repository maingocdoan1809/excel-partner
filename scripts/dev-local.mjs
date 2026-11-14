/**
 * Runs the pane and the bridge from this working copy, with no install or update: the page is rebuilt when a frontend file changes and
 * the bridge restarts when a bridge, shared, providers or protocol file changes. It uses the same data folder, certificate and provider
 * CLIs as the installed product, so the chats and sign-ins you see are the real ones.
 *
 *   npm run dev              port 8765 when the installed bridge is stopped, else 8777
 *   npm run dev -- --port N  a port of your choice
 *
 * Excel, Word and PowerPoint load the pane from port 8765, so to try a change inside Office run `office-partner stop` first.
 * A browser tab at https://localhost:<port>/ works on either port.
 */
import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import fs from "node:fs";
import net from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const store = createRequire(import.meta.url)(path.join(root, "bin", "runtime-store.cjs"));
const dataDir = store.home(process.env);
const cert = path.join(dataDir, "localhost.crt"), key = path.join(dataDir, "localhost.key");
if (!fs.existsSync(cert) || !fs.existsSync(key)) { console.error(`No certificate in ${dataDir}. Run the installer once (office-partner setup) so the pane has a trusted localhost certificate.`); process.exit(1); }

const inUse = port => new Promise(resolve => {
  const socket = net.connect({ port, host: "127.0.0.1" });
  socket.once("connect", () => { socket.destroy(); resolve(true); });
  socket.once("error", () => resolve(false));
});
const asked = process.argv.indexOf("--port");
const port = asked > 0 ? Number(process.argv[asked + 1]) : (await inUse(8765)) ? 8777 : 8765;
if (!Number.isInteger(port) || port < 1 || port > 65535) { console.error("Choose a port from 1 to 65535."); process.exit(1); }
if (await inUse(port)) { console.error(`Port ${port} is already in use. Stop what uses it or pass --port.`); process.exit(1); }

const bin = name => path.join(root, "node_modules", ".bin", process.platform === "win32" ? `${name}.cmd` : name);
const children = [];
const run = (label, command, args, env = {}) => {
  const child = spawn(command, args, { cwd: root, env: { ...process.env, ...env }, stdio: "inherit", shell: process.platform === "win32" && command.endsWith(".cmd") });
  child.on("exit", code => { console.log(`[${label}] stopped${code ? ` (exit ${code})` : ""}`); stopAll(); });
  children.push(child);
};
let stopping = false;
function stopAll() {
  if (stopping) return; stopping = true;
  // `tsx watch` starts the bridge as a child of its own, so on Windows the whole tree is stopped, not just the first process.
  for (const child of children) {
    try { if (process.platform === "win32" && child.pid) spawn("taskkill", ["/pid", String(child.pid), "/T", "/F"], { stdio: "ignore" }); else child.kill(); } catch { /* already gone */ }
  }
  setTimeout(() => process.exit(0), 800);
}
process.on("SIGINT", stopAll); process.on("SIGTERM", stopAll);

run("pane", bin("esbuild"), ["frontend/taskpane/main.ts", "--bundle", "--format=iife", "--target=es2020", "--outfile=frontend/dist/taskpane.js", "--watch=forever", "--log-level=warning"]);
run("bridge", bin("tsx"), ["watch", "--clear-screen=false", "bridge/src/index.ts"], { BRIDGE_PORT: String(port), BRIDGE_CERT: cert, BRIDGE_KEY: key });
console.log(`Dev bridge on https://localhost:${port}/  (data folder ${dataDir}). Press Ctrl+C to stop.`);
if (port !== 8765) console.log("Office still talks to port 8765. Run `office-partner stop`, then `npm run dev` again to use this build inside Excel.");
