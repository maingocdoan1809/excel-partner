import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { killTrackedProcesses, reapRecordedProcesses, setProcessRegistryFile } from "../../providers/src/process-registry.js";
import { crashSummary, CRASH_EXIT_CODE } from "../../shared/src/crash-exit.js";
import { dataDir, loadConfig } from "./config.js";
import { GatewayServer } from "./server.js";

const log = (entry: object): void => void process.stdout.write(JSON.stringify({ at: new Date().toISOString(), ...entry }) + "\n");
const dir = dataDir();
let server: GatewayServer | undefined;

// An unexpected error leaves the gateway in an unknown state: stop the provider CLIs and exit with the crash code so the
// supervisor starts a clean one.
let crashed = false;
const crash = (kind: string, error: unknown): void => {
  if (crashed) return;
  crashed = true;
  log({ event: "crash", kind, ...crashSummary(error) });
  killTrackedProcesses();
  const closing = server ? server.close().catch(() => undefined) : Promise.resolve();
  void Promise.race([closing, new Promise<void>(resolve => setTimeout(resolve, 3000).unref())]).finally(() => process.exit(CRASH_EXIT_CODE));
};
process.on("uncaughtException", error => crash("uncaught_exception", error));
process.on("unhandledRejection", reason => crash("unhandled_rejection", reason));

// Provider CLIs a crashed or killed gateway left running are stopped before any new one starts.
const registry = path.join(dir, "provider-processes.json");
const stopped = reapRecordedProcesses(registry);
if (stopped) log({ event: "stale_provider_processes_stopped", count: stopped });
setProcessRegistryFile(registry);

const stop = () => { void server!.close().finally(() => { killTrackedProcesses(); process.exit(0); }); };
const token = fs.readFileSync(path.join(dir, "token.txt"), "utf8").trim();
const bundledRelay = fileURLToPath(new URL("./relay.mjs", import.meta.url));
server = new GatewayServer({ config: loadConfig(dir), token,
  conversationFile: path.join(dir, "conversations.json"),
  tls: { cert: path.join(dir, "localhost.crt"), key: path.join(dir, "localhost.key") },
  relayEntry: process.env.AGENT_GATEWAY_RELAY_ENTRY ?? (fs.existsSync(bundledRelay) ? bundledRelay : fileURLToPath(new URL("./relay.js", import.meta.url))),
  log,
  onShutdown: stop });
await server.listen();
console.log(`Agent Gateway ready at https://localhost:${loadConfig(dir).port}`);
for (const signal of ["SIGTERM", "SIGINT"] as const) process.on(signal, stop);
