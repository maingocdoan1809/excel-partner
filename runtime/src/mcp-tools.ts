import type { ToolResult } from "../../protocol/src/index.js";
import type { McpServers } from "./mcp-servers.js";

export interface McpToolContext {
  sessionId: string;
  enabled: ReadonlySet<string>;
  signal?: AbortSignal;
  isActive(): boolean;
  authorize(detail: { server: string; tool: string; arguments: Record<string, unknown> }): Promise<boolean>;
}

/** Shared authorization and execution for an actually registered external tool. */
export async function invokeMcpCall(manager: McpServers, context: McpToolContext, serverId: string, toolName: string, args: unknown): Promise<ToolResult> {
  const failure = (message: string, code: "cancelled" | "invalid_arguments" | "internal_error" = "cancelled"): ToolResult => ({ ok: false, error: { code, message } });
  if (!args || typeof args !== "object" || Array.isArray(args)) return failure("MCP arguments must be an object", "invalid_arguments");
  const input = args as Record<string, unknown>;
  const allowed = () => context.enabled.has(serverId) && context.isActive() && !context.signal?.aborted;
  if (!allowed()) return failure("MCP server is not enabled or the request ended");
  try {
    const server = manager.list().find(item => item.id === serverId);
    if (!server) return failure("MCP server not found", "invalid_arguments");
    if (!await context.authorize({ server: server.name, tool: toolName, arguments: input }) || !allowed()) return failure("MCP call was not approved or the request ended");
    return { ok: true, data: { mcpResult: await manager.call(context.sessionId, serverId, toolName, input, context.signal) } };
  } catch (error) { return failure(error instanceof Error ? error.message : String(error), context.signal?.aborted ? "cancelled" : "internal_error"); }
}
