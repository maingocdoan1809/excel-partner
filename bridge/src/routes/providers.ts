import type http from "node:http";
import { ProviderIdSchema, providerSessionKey } from "../../../protocol/src/index.js";
import { cleanSessionTitle } from "../../../protocol/src/session-title.js";
import { normalizeCommands } from "../../../providers/src/commands.js";
import type { ProviderRegistry } from "../../../providers/src/registry.js";
import { connectedSession, sendJson } from "../http-util.js";
import type { Logger } from "../logger.js";
import { ProviderLoginError, type ProviderLoginManager } from "../provider-login.js";
import type { SessionManager } from "../session-manager.js";
import type { UsageStore } from "../usage.js";

export interface ProviderRouteDeps {
  sessions: SessionManager;
  providers: ProviderRegistry;
  logins: ProviderLoginManager;
  usage: UsageStore;
  logger: Logger;
}

/** What the pane asks about the provider CLIs: sign-in, models, slash commands and recorded usage. */
export class ProviderRoutes {
  constructor(private readonly deps: ProviderRouteDeps) {}

  async login(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const { logins, providers } = this.deps;
    if (req.method === "GET") {
      const status = logins.status(session.id);
      if (status.state === "success") providers.invalidateModels("copilot");
      return sendJson(res, 200, status);
    }
    if (req.method === "DELETE") return sendJson(res, 200, logins.cancel(session.id));
    if (req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
    if (session.isBusy) return sendJson(res, 409, { error: "Stop the current response before signing in." });
    try {
      return sendJson(res, 200, logins.start(session.id));
    } catch (error) {
      if (error instanceof ProviderLoginError) return sendJson(res, 409, { error: error.message });
      throw error;
    }
  }

  async models(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    const parsed = ProviderIdSchema.safeParse(url.searchParams.get("provider"));
    if (!parsed.success) return sendJson(res, 400, { error: "invalid_provider" });
    try {
      return sendJson(res, 200, await this.deps.providers.listModels(parsed.data));
    } catch (error) {
      this.deps.logger.warn("could not list provider models", { provider: parsed.data, error: String(error) });
      if (/not (?:authenticated|logged in)|(?:authentication|login|sign.?in) (?:required|failed)|unauthorized|401|please (?:log|sign) in/i.test(String(error))) {
        return sendJson(res, 401, { error: "Sign in to your AI account to load the available models.", code: "login_required" });
      }
      return sendJson(res, 503, { error: "Could not load models from the provider CLI." });
    }
  }

  /** The provider's own slash commands, plus the ones the pane handles itself, for the `/` menu. */
  async commands(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    const provider = ProviderIdSchema.safeParse(url.searchParams.get("provider"));
    if (!provider.success) return sendJson(res, 400, { error: "Invalid provider" });
    try {
      const adapter = this.deps.providers.require(provider.data);
      const commands = await adapter.listCommands?.(session.provider === provider.data ? session.providerSessionId ?? undefined : undefined) ?? [];
      const local = normalizeCommands([
        { name: "new", description: "Start a new chat" }, { name: "clear", description: "Start a new chat" },
        { name: "resume", description: "Browse saved conversations" }, { name: "model", description: "Choose a model" },
        { name: "permissions", description: "Choose Ask or Auto tool approval" }, { name: "theme", description: "Change the add-in theme" },
        { name: "copy", description: "Copy the last assistant response" }, { name: "export", description: "Export this chat as Markdown" },
        { name: "help", description: "Show available commands" }, { name: "login", description: "Open provider sign-in" },
        { name: "status", description: "Show the current provider and session", execution: "client", action: "status" },
        { name: "usage", description: "Show session usage", execution: "client", action: "usage" },
      ]);
      const names = new Set(commands.map(command => command.name));
      return sendJson(res, 200, { provider: provider.data, commands: [...commands, ...local.filter(command => !names.has(command.name))] });
    } catch (error) { return sendJson(res, 502, { error: error instanceof Error ? error.message : String(error) }); }
  }

  async usage(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    const id = url.searchParams.get("sessionId");
    const summaries = this.deps.usage.summaries().filter((item) => !id || item.sessionId === id);
    await Promise.all(summaries.map(async item => {
      const provider = ProviderIdSchema.safeParse(item.provider);
      const nativeId = item.providerSessionIds?.at(-1);
      if (!provider.success || !nativeId || item.sessionId !== providerSessionKey(provider.data, nativeId)) return;
      try {
        item.title = await this.deps.providers.get(provider.data)?.sessionTitle?.(nativeId) ?? cleanSessionTitle(provider.data, item.title);
      } catch { /* Preserve the last recorded title when the CLI is unavailable. */ }
    }));
    return sendJson(res, 200, { sessions: summaries });
  }
}
