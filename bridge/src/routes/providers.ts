import type http from "node:http";
import { DEFAULT_PROVIDER, ProviderIdSchema, providerSessionKey, type ProviderId } from "../../../protocol/src/index.js";
import { cleanSessionTitle } from "../../../protocol/src/session-title.js";
import { normalizeCommands } from "../../../providers/src/commands.js";
import type { ProviderRegistry } from "../../../providers/src/registry.js";
import { connectedSession, readBody, sendJson } from "../http-util.js";
import type { Logger } from "../logger.js";
import type { LoginChoice } from "../../../protocol/src/provider-login.js";
import { ProviderLoginError, type ProviderLoginManager } from "../provider-login.js";
import type { WorkbookSessions } from "../workbook-sessions.js";
import type { UsageStore } from "../usage.js";

export interface ProviderRouteDeps {
  sessions: WorkbookSessions;
  providers: ProviderRegistry;
  logins: ProviderLoginManager;
  usage: UsageStore;
  logger: Logger;
  /** Providers the admin policy turns off: no sign-in, models or commands for them. */
  disabled?: readonly ProviderId[];
}

const DISABLED = { error: "This AI provider is turned off by your administrator.", code: "provider_disabled" };

/** What the pane asks about the provider CLIs: sign-in, models, slash commands and recorded usage. */
export class ProviderRoutes {
  constructor(private readonly deps: ProviderRouteDeps) {}

  async login(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    const { logins, providers } = this.deps;
    const query = new URL(req.url ?? "/", "http://localhost").searchParams;
    const parsed = ProviderIdSchema.safeParse(query.get("provider") ?? DEFAULT_PROVIDER);
    if (!parsed.success) return sendJson(res, 400, { error: "invalid_provider" });
    const provider = parsed.data;
    if (this.deps.disabled?.includes(provider)) return sendJson(res, 403, DISABLED);
    if (req.method === "GET" && query.get("accounts") === "1") {
      try {
        return sendJson(res, 200, { accounts: await logins.accounts(provider) });
      } catch (error) {
        this.deps.logger.warn("could not list sign-in accounts", { provider, error: String(error) });
        return sendJson(res, 503, { error: error instanceof ProviderLoginError ? error.message : "Could not load the accounts. Try again." });
      }
    }
    if (req.method === "GET") {
      const status = logins.status(session.id, provider, query.get("fresh") === "1");
      if (status.state === "success") providers.invalidateModels(provider);
      return sendJson(res, 200, status);
    }
    if (req.method === "DELETE") return sendJson(res, 200, logins.cancel(session.id));
    if (req.method !== "POST" && req.method !== "PUT") return sendJson(res, 405, { error: "method_not_allowed" });
    if (session.isBusy) return sendJson(res, 409, { error: "Stop the current response before signing in." });
    try {
      if (req.method === "PUT") {
        const body = JSON.parse(await readBody(req, 4096) || "{}") as { code?: unknown };
        return sendJson(res, 200, logins.submitCode(session.id, typeof body.code === "string" ? body.code : ""));
      }
      const body = JSON.parse(await readBody(req, 16_384) || "{}") as Partial<LoginChoice>;
      return sendJson(res, 200, logins.start(session.id, provider, typeof body.accountId === "string" ? {
        accountId: body.accountId, method: typeof body.method === "number" ? body.method : -1,
        ...(body.inputs && typeof body.inputs === "object" ? { inputs: body.inputs } : {}), ...(typeof body.key === "string" ? { key: body.key } : {}),
      } : undefined));
    } catch (error) {
      if (error instanceof ProviderLoginError) return sendJson(res, 409, { error: error.message });
      if (error instanceof SyntaxError) return sendJson(res, 400, { error: "invalid_json" });
      throw error;
    }
  }

  async logout(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    if (req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
    if (session.isBusy) return sendJson(res, 409, { error: "Stop the current response before signing out." });
    const parsed = ProviderIdSchema.safeParse(new URL(req.url ?? "/", "http://localhost").searchParams.get("provider"));
    if (!parsed.success) return sendJson(res, 400, { error: "invalid_provider" });
    if (this.deps.disabled?.includes(parsed.data)) return sendJson(res, 403, DISABLED);
    try {
      const body = JSON.parse(await readBody(req, 4096) || "{}") as { accountId?: unknown };
      await this.deps.logins.logout(parsed.data, typeof body.accountId === "string" ? body.accountId : undefined);
      this.deps.providers.invalidateModels(parsed.data);
      return sendJson(res, 200, { ok: true });
    } catch (error) {
      if (error instanceof ProviderLoginError) return sendJson(res, 409, { error: error.message });
      if (error instanceof SyntaxError) return sendJson(res, 400, { error: "invalid_json" });
      throw error;
    }
  }

  /** Changes one of the settings the CLI offers for this chat's open session (see ProviderAdapter.setSessionOption). */
  async option(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const session = connectedSession(this.deps.sessions, req);
    if (!session) return sendJson(res, 401, { error: "unauthorized" });
    res.setHeader("cache-control", "no-store");
    if (req.method !== "POST") return sendJson(res, 405, { error: "method_not_allowed" });
    if (session.isBusy) return sendJson(res, 409, { error: "Wait for the current response to finish before changing this." });
    const adapter = session.provider ? this.deps.providers.get(session.provider) : undefined;
    if (!adapter?.setSessionOption) return sendJson(res, 400, { error: "This provider has no settings to change." });
    if (!session.providerSessionId) return sendJson(res, 409, { error: "Send a message first: the setting applies to the AI's open session." });
    try {
      const body = JSON.parse(await readBody(req, 4096) || "{}") as { id?: unknown; value?: unknown };
      if (typeof body.id !== "string" || typeof body.value !== "string") return sendJson(res, 400, { error: "Specify the setting and its value." });
      return sendJson(res, 200, { options: await adapter.setSessionOption(session.providerSessionId, body.id, body.value) });
    } catch (error) {
      if (error instanceof SyntaxError) return sendJson(res, 400, { error: "invalid_json" });
      return sendJson(res, 409, { error: error instanceof Error ? error.message : "Could not change the setting." });
    }
  }

  async models(req: http.IncomingMessage, res: http.ServerResponse, url: URL): Promise<void> {
    if (!connectedSession(this.deps.sessions, req)) return sendJson(res, 401, { error: "unauthorized" });
    const parsed = ProviderIdSchema.safeParse(url.searchParams.get("provider"));
    if (!parsed.success) return sendJson(res, 400, { error: "invalid_provider" });
    if (this.deps.disabled?.includes(parsed.data)) return sendJson(res, 403, DISABLED);
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
    if (this.deps.disabled?.includes(provider.data)) return sendJson(res, 403, DISABLED);
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
    // Someone is looking at the usage: now is when a CLI that does not report it turn by turn is asked for its totals.
    const resolve = url.searchParams.get("resolve") === "1";
    if (resolve) {
      await Promise.all(this.deps.usage.unsettled(id).map(async ({ provider, providerSessionId }) => {
        const parsed = ProviderIdSchema.safeParse(provider);
        const adapter = parsed.success ? this.deps.providers.get(parsed.data) : undefined;
        if (!adapter?.sessionUsage) return;
        const asked = Date.now();
        try { this.deps.usage.settle(provider, providerSessionId, (await adapter.sessionUsage(providerSessionId)).usage, asked); }
        catch (error) { this.deps.logger.warn("could not read the CLI's usage", { provider, error: String(error) }); }
      }));
    }
    const summaries = this.deps.usage.summaries().filter((item) => !id || item.sessionId === id);
    await Promise.all(summaries.map(async item => {
      const provider = ProviderIdSchema.safeParse(item.provider);
      const nativeId = item.providerSessionIds?.at(-1);
      if (!provider.success || !nativeId || item.sessionId !== providerSessionKey(provider.data, nativeId)) return;
      try {
        item.title = await this.deps.providers.get(provider.data)?.sessionTitle?.(nativeId, { cheapOnly: !resolve }) ?? cleanSessionTitle(item.title);
      } catch { /* Preserve the last recorded title when the CLI is unavailable. */ }
    }));
    return sendJson(res, 200, { sessions: summaries });
  }
}
