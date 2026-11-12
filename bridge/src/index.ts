#!/usr/bin/env node
import { loadConfig } from "./config.js";
import { Logger } from "./logger.js";
import { Bridge } from "./server.js";
import { stopWorkspaceHelper, warmWorkspaceHelper } from "./windows-workspace.js";

const SHUTDOWN_GRACE_MS = 8000;

async function main(): Promise<void> {
  const config = loadConfig();
  const logger = new Logger("bridge", config.logLevel);
  const bridge = new Bridge(config, { logger, onShutdown: () => { void shutdown("management"); } });

  const { url } = await bridge.listen();
  logger.info("ready", { url, health: `${url}/health` });
  warmWorkspaceHelper();

  let shuttingDown = false;
  const shutdown = async (signal: string): Promise<void> => {
    if (shuttingDown) return;
    shuttingDown = true;
    logger.info("shutting down", { signal });
    try {
      // A close that hangs (a provider worker that will not stop, a socket that stays open) must not leave a half-closed bridge
      // running: its history database is already shut and it keeps refusing requests, yet it still answers /health.
      await Promise.race([bridge.close(), new Promise<void>((resolve) => setTimeout(resolve, SHUTDOWN_GRACE_MS).unref())]);
    } finally {
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
