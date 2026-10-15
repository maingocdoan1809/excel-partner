#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { HttpToolForwarder } from "./forwarder.js";
import { createExcelMcpServer } from "./server.js";

/**
 * Entry point launched by provider CLIs as their MCP server over stdio.
 *
 * Required environment (injected by the bridge when it starts a provider
 * session):
 *   EXCEL_BRIDGE_URL      e.g. https://127.0.0.1:8765
 *   EXCEL_SESSION_ID      workbook session id
 *   EXCEL_SESSION_TOKEN   per-session bearer token
 */
async function main(): Promise<void> {
  const bridgeUrl = process.env.EXCEL_BRIDGE_URL;
  const sessionId = process.env.EXCEL_SESSION_ID;
  const token = process.env.EXCEL_SESSION_TOKEN;

  if (!bridgeUrl || !sessionId || !token) {
    process.stderr.write(
      "excel-mcp: missing EXCEL_BRIDGE_URL / EXCEL_SESSION_ID / EXCEL_SESSION_TOKEN\n",
    );
    process.exit(2);
  }

  const forwarder = new HttpToolForwarder({ bridgeUrl, sessionId, token });
  const server = createExcelMcpServer(forwarder, { allowOptInTools: process.env.EXCEL_ALLOW_OFFICE_JS === "1" });
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  process.stderr.write(`excel-mcp fatal: ${error instanceof Error ? error.stack : error}\n`);
  process.exit(1);
});
