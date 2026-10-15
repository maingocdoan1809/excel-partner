#!/usr/bin/env node
import { loadConfig } from "./config.js";
import { Logger } from "./logger.js";
import { Bridge } from "./server.js";
import { stopWorkspaceHelper, warmWorkspaceHelper } from "./windows-workspace.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const logger = new Logger("bridge", config.logLevel);
  const bridge = new Bridge(config, { logger });

  const { url } = await bridge.listen();
  logger.info("ready", { url, health: `${url}/health` });
  warmWorkspaceHelper();

  let shuttingDown = false;
  const shutdown = async (signal: string): Promise<void> => {
    if (shuttingDown) return;
    shuttingDown = true;
    logger.info("shutting down", { signal });
    try {
      await bridge.close();
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
