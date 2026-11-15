import type http from "node:http";
import { z } from "zod";
import { ProviderIdSchema, providerSessionKey, type ProviderId } from "../../../protocol/src/index.js";
import type { StoredSession } from "../../../protocol/src/history.js";
import type { ProviderRegistry } from "../../../providers/src/registry.js";
import { HistoryError } from "../history-store.js";
import { connectedSession, readBody, sendJson } from "../http-util.js";
import type { Logger } from "../logger.js";
import type { WorkbookSessions } from "../workbook-sessions.js";
import type { HistoryRoutes } from "./history.js";

export interface ExternalSessionDeps {
  sessions: WorkbookSessions;
  providers: ProviderRegistry;
  history: HistoryRoutes;
  logger: Logger;
  /** Providers the admin policy turns off. */
  disabled?: readonly ProviderId[];
}

const ImportBody = z.object({ provider: ProviderIdSchema, id: z.string().min(1).max(256) });

/** Conversations a provider's own tools (its terminal CLI, its VS Code extension) saved: what there is to import, and the import. */
export class ExternalSessionRoutes {
  constructor(private readonly deps: ExternalSessionDeps) {}

  async handle(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    if (req.method !== "GET" && req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
    try {
      const body = req.method === "POST" ? ImportBody.parse(JSON.parse(await readBody(req, 4096))) : null;
      const parsed = ProviderIdSchema.safeParse(body?.provider ?? url.searchParams.get("provider"));
      if (!parsed.success) return sendJson(res, 400, { error: "invalid_provider" });
      if (this.deps.disabled?.includes(parsed.data)) return sendJson(res, 403, { error: "This AI provider is turned off by your administrator.", code: "provider_disabled" });
      const adapter = this.deps.providers.get(parsed.data);
      if (!adapter?.capabilities.importSessions || !adapter.listExternalSessions || !adapter.importExternalSession) {
        return sendJson(res, 400, { error: "This provider cannot import conversations." });
      }
      if (!body) return sendJson(res, 200, { sessions: await adapter.listExternalSessions() });
      const imported = await adapter.importExternalSession(body.id);
      // The id every chat of a native session has, so continuing it here carries on the same History item.
      const chat: StoredSession = {
        id: providerSessionKey(parsed.data, imported.providerSessionId), title: imported.title, provider: parsed.data, host: session.host,
        providerSessionId: imported.providerSessionId, identityVersion: 1, startedAt: imported.startedAt, updatedAt: imported.updatedAt,
        messages: imported.messages, changes: [],
      };
      return sendJson(res, 200, this.deps.history.open().saveImported(session.host, chat));
    } catch (error) {
      if (error instanceof z.ZodError || error instanceof SyntaxError) return sendJson(res, 400, { error: "Invalid request." });
      if (error instanceof HistoryError) return sendJson(res, 503, { code: error.code, error: "History could not be saved. Check disk space and permissions." });
      this.deps.logger.warn("external sessions failed", { error: error instanceof Error ? error.message : String(error) });
      return sendJson(res, 409, { error: error instanceof Error ? error.message : "Could not read the conversations." });
    }
  }
}
