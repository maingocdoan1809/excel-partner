#!/usr/bin/env node
import path from "node:path";
import { loadAdminPolicy } from "./policy.js";
import { killTrackedProcesses, reapRecordedProcesses, setProcessRegistryFile } from "../../providers/src/process-registry.js";
import { loadConfig } from "./config.js";
import { crashSummary, CRASH_EXIT_CODE } from "../../shared/src/crash-exit.js";
import { maintainHistory } from "./history-store.js";
import { Logger } from "./logger.js";
import { Bridge } from "./server.js";
import { stopWorkspaceHelper, warmWorkspaceHelper } from "./windows-workspace.js";

const SHUTDOWN_GRACE_MS = 8000;
const CRASH_GRACE_MS = 3000;
const DAY_MS = 24 * 60 * 60_000;

/** The admin policy's history retention, read again for each run so a changed policy applies without a restart. */
const historyRetentionDays = () => loadAdminPolicy().policy.historyRetentionDays;

async function main(): Promise<void> {
  const config = loadConfig();
  const logger = new Logger("bridge", config.logLevel);
  let bridge: Bridge | undefined;

  // An unexpected error leaves the bridge in an unknown state: stop the provider CLIs, let the history database close,
  // and exit with a code the supervisor recognises as a crash, so it starts a clean bridge.
  let crashed = false;
  const crash = (kind: string, error: unknown): void => {
    if (crashed) return;
    crashed = true;
    logger.error("bridge crashed", { kind, ...crashSummary(error) });
    killTrackedProcesses();
    stopWorkspaceHelper();
    const closing = bridge ? bridge.close().catch(() => undefined) : Promise.resolve();
    void Promise.race([closing, new Promise<void>((resolve) => setTimeout(resolve, CRASH_GRACE_MS).unref())])
      .finally(() => process.exit(CRASH_EXIT_CODE));
  };
  process.on("uncaughtException", (error) => crash("uncaught_exception", error));
  process.on("unhandledRejection", (reason) => crash("unhandled_rejection", reason));

  // Provider CLIs a crashed or killed bridge left running are stopped before any new one starts.
  const registry = path.join(path.dirname(config.configFile), "provider-processes.json");
  const stopped = reapRecordedProcesses(registry);
  if (stopped) logger.warn("stopped provider processes left by a previous bridge", { count: stopped });
  setProcessRegistryFile(registry);

  bridge = new Bridge(config, { logger, onShutdown: () => { void shutdown("management"); } });

  const { url } = await bridge.listen();
  logger.info("ready", { url, health: `${url}/health` });
  warmWorkspaceHelper();

  // History housekeeping (retention policy, daily backup) now and once a day. A failure only skips this run.
  const historyFile = path.join(path.dirname(config.configFile), "history.db");
  const maintain = () => maintainHistory(historyFile, historyRetentionDays())
    .then(({ removed, backedUp }) => { if (removed || backedUp) logger.info("history maintenance", { removed, backedUp }); })
    .catch((error) => logger.warn("history maintenance skipped", { error: error instanceof Error ? error.message : String(error) }));
  void maintain();
  setInterval(() => void maintain(), DAY_MS).unref();

  let shuttingDown = false;
  const shutdown = async (signal: string): Promise<void> => {
    if (shuttingDown) return;
    shuttingDown = true;
    logger.info("shutting down", { signal });
    try {
      // A close that hangs (a provider worker that will not stop, a socket that stays open) must not leave a half-closed bridge
      // running: its history database is already shut and it keeps refusing requests, yet it still answers /health.
      await Promise.race([bridge!.close(), new Promise<void>((resolve) => setTimeout(resolve, SHUTDOWN_GRACE_MS).unref())]);
    } finally {
      killTrackedProcesses();
      stopWorkspaceHelper();
      process.exit(0);
    }
  };

  process.on("SIGINT", () => void shutdown("SIGINT"));
  process.on("SIGTERM", () => void shutdown("SIGTERM"));
}

main().catch((error) => {
  process.stderr.write(`bridge fatal: ${error instanceof Error ? error.stack : error}\n`);
  process.exit(1);
});

export { Bridge } from "./server.js";
export { loadConfig } from "./config.js";
