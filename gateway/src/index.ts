import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { dataDir, loadConfig } from "./config.js";
import { GatewayServer } from "./server.js";

const dir = dataDir();
const token = fs.readFileSync(path.join(dir, "token.txt"), "utf8").trim();
const bundledRelay = fileURLToPath(new URL("./relay.mjs", import.meta.url));
const server = new GatewayServer({ config: loadConfig(dir), token,
  conversationFile: path.join(dir, "conversations.json"),
  tls: { cert: path.join(dir, "localhost.crt"), key: path.join(dir, "localhost.key") },
  relayEntry: process.env.AGENT_GATEWAY_RELAY_ENTRY ?? (fs.existsSync(bundledRelay) ? bundledRelay : fileURLToPath(new URL("./relay.js", import.meta.url))),
  log: entry => process.stdout.write(JSON.stringify({ at: new Date().toISOString(), ...entry }) + "\n"),
  onShutdown: () => { void server.close().then(() => process.exit(0)); } });
await server.listen();
console.log(`Agent Gateway ready at https://localhost:${loadConfig(dir).port}`);
for (const signal of ["SIGTERM", "SIGINT"] as const) process.on(signal, () => { void server.close().then(() => process.exit(0)); });
