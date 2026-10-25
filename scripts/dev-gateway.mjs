// Runs the gateway from source with auto-restart, isolated from an installed gateway.
//   npm run dev:gateway            (add `-- --trust` once on Windows to trust the dev certificate)
import { spawn, spawnSync } from "node:child_process";
import { randomBytes } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { context } from "esbuild";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const home = path.resolve(process.env.AGENT_GATEWAY_HOME ?? path.join(root, "run", "gateway-dev"));
const cert = path.join(home, "localhost.crt"), key = path.join(home, "localhost.key");
const tokenFile = path.join(home, "token.txt"), configFile = path.join(home, "config.json");

fs.mkdirSync(home, { recursive: true });
// Port 8767 and no startup warmup: avoids clashing with an installed gateway and spending quota.
// Compatibility models let the client pick claude-* names that route to the default provider.
if (!fs.existsSync(configFile)) fs.writeFileSync(configFile, JSON.stringify({ port: 8767, warmupOnStart: false, advertiseCompatibilityModels: true }, null, 2), { mode: 0o600 });
if (!fs.existsSync(tokenFile)) fs.writeFileSync(tokenFile, `agw-${randomBytes(32).toString("hex")}`, { mode: 0o600 });
// Prefer a localhost certificate this machine already trusts (installed gateway / Excel Partner).
const local = process.env.LOCALAPPDATA ?? path.join(process.env.HOME ?? "", ".config");
for (const name of ["Agent Gateway", "Excel Partner"]) {
  if (fs.existsSync(cert) && fs.existsSync(key)) break;
  const [c, k] = ["localhost.crt", "localhost.key"].map(file => path.join(local, name, file));
  if (fs.existsSync(c) && fs.existsSync(k)) { fs.copyFileSync(c, cert); fs.copyFileSync(k, key); console.log(`Reusing certificate from ${path.join(local, name)}`); }
}
if (!fs.existsSync(cert) || !fs.existsSync(key)) {
  const { default: selfsigned } = await import("selfsigned");
  const pems = await selfsigned.generate([{ name: "commonName", value: "localhost" }], { algorithm: "sha256", notAfterDate: new Date(Date.now() + 365 * 86400000), extensions: [
    { name: "basicConstraints", cA: false }, { name: "keyUsage", digitalSignature: true, keyEncipherment: true },
    { name: "extKeyUsage", serverAuth: true }, { name: "subjectAltName", altNames: [{ type: 2, value: "localhost" }, { type: 7, ip: "127.0.0.1" }, { type: 7, ip: "::1" }] },
  ] });
  fs.writeFileSync(cert, pems.cert, { mode: 0o600 }); fs.writeFileSync(key, pems.private, { mode: 0o600 });
}
if (process.argv.includes("--trust")) {
  if (process.platform !== "win32") console.log(`Trust this certificate in your OS/browser: ${cert}`);
  else if (spawnSync("certutil.exe", ["-addstore", "-user", "Root", cert], { stdio: "inherit", windowsHide: true }).status !== 0) throw new Error("Could not trust the dev certificate");
}

// The relay is spawned by provider CLIs as plain node, so it must be a built file.
const relay = path.join(root, "gateway", "dist", "relay.mjs");
const watcher = await context({ absWorkingDir: root, entryPoints: { relay: "gateway/src/relay.ts" }, outdir: "gateway/dist", outExtension: { ".js": ".mjs" },
  platform: "node", format: "esm", target: "node20", bundle: true, packages: "external", sourcemap: true });
await watcher.rebuild();
await watcher.watch();

const port = JSON.parse(fs.readFileSync(configFile, "utf8")).port ?? 8766;
console.log(`Agent Gateway (dev)\n  URL:    https://localhost:${port}\n  Token:  ${fs.readFileSync(tokenFile, "utf8").trim()}\n  Home:   ${home}\n`);

const tsx = path.join(root, "node_modules", "tsx", "dist", "cli.mjs");
const child = spawn(process.execPath, [tsx, "watch", "--clear-screen=false", "gateway/src/index.ts"], { cwd: root, stdio: "inherit",
  env: { ...process.env, AGENT_GATEWAY_HOME: home, AGENT_GATEWAY_RELAY_ENTRY: relay } });
const stopChild = () => {
  // tsx watch spawns the server as a grandchild; on Windows kill the whole tree.
  if (process.platform === "win32" && child.pid) spawnSync("taskkill", ["/PID", String(child.pid), "/T", "/F"], { stdio: "ignore", windowsHide: true });
  else child.kill();
};
const shutdown = () => { stopChild(); void watcher.dispose().then(() => process.exit(0)); };
process.on("SIGINT", shutdown); process.on("SIGTERM", shutdown);
process.on("exit", stopChild);
child.on("exit", code => { void watcher.dispose().then(() => process.exit(code ?? 0)); });
