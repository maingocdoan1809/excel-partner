import path from "node:path";
import { terminalAgents } from "../../providers/src/registry.js";
import { syncTerminalAgents, type TerminalMcpServer } from "../../providers/src/terminal-agent.js";
import type { BridgeConfig } from "./config.js";
import type { Logger } from "./logger.js";

/** The MCP server a terminal agent uses to reach the open Office files: the same node and release as the Excel MCP, a sibling entry file. */
export function peerServerFor(config: BridgeConfig): TerminalMcpServer {
  return {
    name: "office-partner",
    command: config.mcp.command,
    args: [path.join(path.dirname(config.mcp.args[0]!), "peer-cli.js")],
    env: {
      OFFICE_PARTNER_DATA_DIR: path.dirname(config.configFile),
      OFFICE_PARTNER_URL: config.publicBaseUrl,
      ...(config.tls ? { OFFICE_PARTNER_CA: config.tls.certPath } : {}),
    },
  };
}

/**
 * Makes every terminal CLI on this computer (Claude Code, Copilot CLI, Codex CLI) able to ask the assistant of an open file,
 * without the person running anything. Runs in the background after start; a CLI that is missing or refuses is only logged.
 * Turn it off with OFFICE_PARTNER_TERMINAL_AGENTS=0 or `{"disabled": true}` in terminal-agents.json in the data folder.
 */
export function registerTerminalAgents(config: BridgeConfig, logger: Logger): void {
  const stateFile = path.join(path.dirname(config.configFile), "terminal-agents.json");
  void syncTerminalAgents(terminalAgents(), peerServerFor(config), { stateFile }).then((results) => {
    for (const [provider, result] of Object.entries(results)) {
      if (result.outcome === "failed") logger.warn("could not register the terminal agent", { provider, detail: result.detail });
      else if (result.outcome === "registered") logger.info("registered Office Partner in the terminal CLI", { provider });
    }
  }, (error) => logger.warn("terminal agent registration failed", { error: String(error) }));
}
