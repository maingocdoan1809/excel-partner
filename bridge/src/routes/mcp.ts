import type http from "node:http";
import type { McpServers } from "../../../runtime/src/mcp-servers.js";
import { connectedSession, readBody, sendJson } from "../http-util.js";
import type { SessionManager } from "../session-manager.js";
import type { WorkbookSession } from "../session.js";

export interface McpRouteDeps {
  sessions: SessionManager;
  mcpServers: McpServers;
  /** Discovers the enabled servers' tools for a pane and records them on its session. */
  discover: (session: WorkbookSession) => Promise<unknown>;
}

/** `/session/mcp`: the external MCP servers Settings manages and the per-pane on/off selection. */
export class McpRoutes {
  constructor(private readonly deps: McpRouteDeps) {}

  async handle(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const { sessions, mcpServers } = this.deps;
    const session = connectedSession(sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    try {
      if (req.method === "GET") return sendJson(res, 200, { servers: mcpServers.list(), enabled: [...session.enabledMcpServers] });
      if (req.method !== "POST" && req.method !== "DELETE") return sendJson(res, 405, { error: "method_not_allowed" });
      const body = JSON.parse(await readBody(req, 32_000)) as Record<string, unknown>;
      const id = typeof body.id === "string" ? body.id : "";
      if (req.method === "DELETE") {
        await mcpServers.remove(id);
        for (const pane of sessions.list()) { pane.enabledMcpServers.delete(id); pane.externalMcpTools.clear(); }
      } else if (body.action === "select") {
        if (typeof body.enabled !== "boolean" || !mcpServers.list().some(item => item.id === id)) throw new Error("Invalid MCP selection");
        if (body.enabled) {
          const alreadyEnabled = session.enabledMcpServers.has(id);
          const previous = session.externalMcpTools;
          session.enabledMcpServers.add(id);
          try { await this.deps.discover(session); }
          catch (error) { if (!alreadyEnabled) session.enabledMcpServers.delete(id); session.externalMcpTools = previous; throw error; }
        } else {
          session.enabledMcpServers.delete(id);
          for (const [name, tool] of session.externalMcpTools) if (tool.serverId === id) session.externalMcpTools.delete(name);
          await mcpServers.closeSessionServer(session.id, id);
        }
      } else if (body.action === "test") {
        const testSession = `test-${session.id}`;
        try { const tools = await mcpServers.tools(testSession, id); return sendJson(res, 200, { tools: tools.map(({ name, description }) => ({ name, description })) }); }
        finally { await mcpServers.closeSession(testSession); }
      } else if (body.action === "save") {
        await mcpServers.save(body.config);
        for (const pane of sessions.list()) { pane.enabledMcpServers.delete((body.config as { id: string }).id); pane.externalMcpTools.clear(); }
      } else throw new Error("Unknown MCP action");
      return sendJson(res, 200, { ok: true });
    } catch (error) { sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) }); }
  }
}
