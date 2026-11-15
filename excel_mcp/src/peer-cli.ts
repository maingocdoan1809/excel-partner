#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createPeerServer, localPeerClient } from "./peer.js";

/**
 * MCP server over stdio for a terminal agent, started as `office-partner mcp-peer`. Register it once, for example
 * `claude mcp add office-partner -- office-partner mcp-peer`. It needs no environment: it finds the bridge, its secret and
 * its certificate where Office Partner keeps them (OFFICE_PARTNER_URL / OFFICE_PARTNER_DATA_DIR override).
 */
createPeerServer(localPeerClient()).connect(new StdioServerTransport()).catch((error) => {
  process.stderr.write(`office-partner mcp-peer fatal: ${error instanceof Error ? error.stack : error}\n`);
  process.exit(1);
});
