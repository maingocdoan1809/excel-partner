import { createHash } from "node:crypto";
import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import type { McpServers } from "./mcp-servers.js";
import { invokeMcpCall, type McpToolContext } from "./mcp-tools.js";

export interface RegisteredMcpTool { serverId: string; toolName: string; definition: Tool }
export type ExternalMcpCatalog = Map<string, RegisteredMcpTool>;
function registeredName(serverId: string, tool: string): string {
  const server = createHash("sha256").update(serverId).digest("hex").slice(0, 10);
  const safe = tool.replace(/[^A-Za-z0-9_-]/g, "_").slice(0, 80);
  const suffix = safe === tool ? "" : `_${createHash("sha256").update(tool).digest("hex").slice(0, 10)}`;
  return `ext_${server}__${safe}${suffix}`;
}
/** Expose actual server tools and schemas, with stable server namespaces to avoid collisions. */
export async function discoverMcpCatalog(manager: McpServers, sessionId: string, enabled: ReadonlySet<string>, reserved: ReadonlySet<string> = new Set(), signal?: AbortSignal): Promise<ExternalMcpCatalog> {
  const catalog: ExternalMcpCatalog = new Map();
  for (const server of manager.list()) {
    if (!enabled.has(server.id)) continue;
    for (const tool of await manager.tools(sessionId, server.id, signal)) {
      const name = registeredName(server.id, tool.name);
      if (reserved.has(name) || catalog.has(name)) throw new Error(`Duplicate registered MCP tool: ${name}`);
      catalog.set(name, { serverId: server.id, toolName: tool.name, definition: { ...tool, name, description: `[${server.name}] ${tool.description ?? tool.name}` } });
    }
  }
  return catalog;
}
export function invokeRegisteredMcpTool(manager: McpServers, context: McpToolContext, tool: RegisteredMcpTool, args: unknown) {
  return invokeMcpCall(manager, context, tool.serverId, tool.toolName, args);
}
