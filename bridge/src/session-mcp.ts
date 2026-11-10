import type { McpEndpoint } from "../../providers/src/types.js";
import type { BridgeConfig } from "./config.js";
import { officeToolCatalog } from "../../runtime/src/catalog.js";
import { discoverMcpCatalog } from "../../runtime/src/mcp-catalog.js";
import type { McpServers } from "../../runtime/src/mcp-servers.js";
import type { WorkbookSession } from "./session.js";

/**
 * Discover the tools of the external MCP servers a pane has enabled and remember them on the session. The Office
 * catalog is passed along so a server cannot shadow a built-in tool name.
 */
export async function discoverSessionMcp(mcpServers: McpServers, allowOfficeJs: boolean, session: WorkbookSession) {
  const enabled = JSON.stringify([...session.enabledMcpServers].sort());
  const catalog = await discoverMcpCatalog(mcpServers, session.id, session.enabledMcpServers,
    new Set(officeToolCatalog(session.host, allowOfficeJs, session.capabilities).map(tool => tool.name)), session.turnController?.signal);
  if (enabled !== JSON.stringify([...session.enabledMcpServers].sort())) throw new Error("MCP selection changed during discovery; retry the request");
  session.externalMcpTools = catalog;
  return catalog;
}

/** Provider MCP transport and native tool names for this pane. */
export function sessionMcpEndpoint(config: BridgeConfig, session: WorkbookSession): McpEndpoint {
  return {
    bridgeUrl: config.publicBaseUrl,
    sessionId: session.id,
    token: session.sessionToken,
    ...(session.provider === "copilot" && session.externalMcpTools.size ? { toolNames: [...officeToolCatalog(session.host, config.allowOfficeJs, session.capabilities).map(tool => tool.name), ...session.externalMcpTools.keys()] } : {}),
    command: config.mcp.command,
    args: config.mcp.args,
    env: {
      EXCEL_BRIDGE_URL: config.publicBaseUrl,
      EXCEL_SESSION_ID: session.id,
      EXCEL_SESSION_TOKEN: session.sessionToken,
      OFFICE_HOST: session.host,
      ...(session.enabledMcpServers.size ? { EXCEL_DYNAMIC_TOOLS: "1" } : {}),
      ...(session.capabilities.webSearch ? { EXCEL_WEB_SEARCH: "1" } : {}),
      ...(session.capabilities.playwright ? { EXCEL_PLAYWRIGHT: "1" } : {}),
      ...(config.allowOfficeJs ? { EXCEL_ALLOW_OFFICE_JS: "1" } : {}),
      ...(config.tls ? { NODE_EXTRA_CA_CERTS: config.tls.certPath } : {}),
    },
  };
}
