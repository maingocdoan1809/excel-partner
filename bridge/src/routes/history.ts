import type http from "node:http";
import { z } from "zod";
import { HistorySaveSchema, StoredSessionSchema, type StoredSession } from "../../../protocol/src/history.js";
import { HistoryStore, HistoryError } from "../history-store.js";
import type { SessionManager } from "../session-manager.js";
import type { Logger } from "../logger.js";
import { connectedSession, readBody, sendJson } from "../http-util.js";
export class HistoryRoutes {
  private store?: HistoryStore;
  constructor(private readonly sessions: SessionManager, private readonly file: string, private readonly logger?: Logger) {}
  close() { this.store?.close(); }
  async handle(req: http.IncomingMessage, res: http.ServerResponse, url: URL) {
    const session = connectedSession(this.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    try {
      const store = this.store ??= new HistoryStore(this.file), host = session.host;
      if (req.method === "GET") {
        const view = url.searchParams.get("view") ?? "list";
        const offset = z.coerce.number().int().min(0).parse(url.searchParams.get("offset") ?? 0);
        const limit = z.coerce.number().int().min(1).max(200).parse(url.searchParams.get("limit") ?? 40);
        if (view === "stats") return sendJson(res, 200, store.stats(host));
        if (view === "export") return sendJson(res, 200, { schemaVersion: 1, host, sessions: store.export(host) });
        if (view === "chat") {
          const record = store.page(host, z.string().min(1).max(256).parse(url.searchParams.get("id")), offset, limit);
          return record ? sendJson(res, 200, record) : sendJson(res, 404, { code: "HISTORY_NOT_FOUND", error: "This chat is unavailable or was deleted." });
        }
        return sendJson(res, 200, store.list(host, offset, limit));
      }
      if (req.method === "DELETE") {
        const id = url.searchParams.get("id");
        if (id) store.delete(host, z.string().min(1).max(256).parse(id)); else store.clear(host);
        return sendJson(res, 200, store.stats(host));
      }
      if (req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
      const body = JSON.parse(await readBody(req, 32 * 1024 * 1024));
      if (body.action === "save") return sendJson(res, 200, store.save(host, HistorySaveSchema.parse(body) as never));
      if (body.action === "import") {
        const records = z.array(StoredSessionSchema).max(100).parse(body.sessions) as StoredSession[];
        // Legacy input is shared in localStorage; each app migrates only its own chats.
        const imported = store.importLegacy(records.filter(item => (item.host ?? "excel") === host));
        return sendJson(res, 200, { imported });
      }
      if (body.action === "cleanup") {
        const before = z.number().finite().positive().parse(body.before);
        const protectedIds = z.array(z.string().max(256)).max(500).parse(body.protectedIds ?? []);
        for (const active of this.sessions.list()) if (active.host === host && active.providerChatSessionId) protectedIds.push(active.providerChatSessionId);
        const removed = store.cleanup(host, before, protectedIds);
        return sendJson(res, 200, { removed, ...store.stats(host) });
      }
      return sendJson(res, 400, { error: "Unknown history action." });
    } catch (error) {
      const tooLarge = error instanceof Error && error.message === "body too large";
      const code = error instanceof HistoryError ? error.code : tooLarge ? "HISTORY_TOO_LARGE" : error instanceof z.ZodError || error instanceof SyntaxError ? "HISTORY_INVALID_INPUT" : "HISTORY_STORAGE_FAILED";
      if (code === "HISTORY_STORAGE_FAILED" || code === "HISTORY_DRIVER_UNAVAILABLE") this.logger?.warn("history_unavailable", { code, error: error instanceof Error ? error.message : String(error) });
      const conflict = ["HISTORY_CONFLICT", "HISTORY_DELETED", "HISTORY_GENERATION"].includes(code);
      if (conflict) return sendJson(res, 409, { code, error: "History changed in another window or was deleted. Your pending copy is retained; export it before resolving the conflict." });
      if (code === "HISTORY_TOO_LARGE") return sendJson(res, 413, { code, error: "This chat is too large to save. Export what you need, then start a new chat." });
      if (code === "HISTORY_INVALID_INPUT") return sendJson(res, 400, { code, error: "History could not be saved or loaded. Check disk space and permissions. Existing data was not replaced." });
      if (code === "HISTORY_DRIVER_UNAVAILABLE") return sendJson(res, 503, { code, error: "History storage is not available with this Node.js version. Run office-partner doctor and install a supported Node.js release. Pending chats are kept in this pane." });
      return sendJson(res, 503, { code, error: "History could not be saved or loaded. Check disk space and permissions. Existing data was not replaced." });
    }
  }
}
