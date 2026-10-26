import http from "node:http";
import https from "node:https";
import fs from "node:fs";
import path from "node:path";
import { randomBytes, randomUUID, createHash, timingSafeEqual } from "node:crypto";
import { fileURLToPath } from "node:url";
import { createDefaultRegistry, type ProviderRegistry } from "../../providers/src/registry.js";
import { GatewayEngine } from "./engine.js";
import { ConversationStore } from "./conversations.js";
import { ApiError, parseRequest, stable, type ApiMessage, type MessagesRequest } from "./protocol.js";
import { publicModelId, routeModel, type GatewayConfig, type Route } from "./config.js";

function matches(value: string | undefined, secret: string) {
  if (!value) return false;
  const a = Buffer.from(value), b = Buffer.from(secret);
  return a.length === b.length && timingSafeEqual(a, b);
}
export interface ServerOptions {
  config: GatewayConfig; token: string; tls: { cert: string; key: string } | null;
  registry?: ProviderRegistry; relayEntry?: string; log?: (entry: object) => void;
  onShutdown?: () => void;
  conversationFile?: string;
}
export class GatewayServer {
  private readonly servers: Array<http.Server | https.Server> = [];
  private readonly relayToken = randomBytes(32).toString("hex");
  private readonly registry: ProviderRegistry;
  private engine!: GatewayEngine;
  private url = "";
  private stopping = false;
  private ready = false;
  private readonly probes = new Map<string, { message: ApiMessage; expires: number }>();
  private readonly probing = new Map<string, Promise<ApiMessage>>();
  private catalog?: { value: Awaited<ReturnType<GatewayServer["computeModels"]>>; expires: number };
  private catalogLoading?: Promise<Awaited<ReturnType<GatewayServer["computeModels"]>>>;
  constructor(readonly options: ServerOptions) {
    if (!options.token || options.token.length < 16) throw new Error("Gateway token must contain at least 16 characters");
    this.registry = options.registry ?? createDefaultRegistry({ bridgeUrl: "", mcp: { command: "", args: [] },
      logger: { debug() {}, warn: message => options.log?.({ event: "provider_warning", message }) } });
  }
  async listen() {
    const handler = (req: http.IncomingMessage, res: http.ServerResponse) => {
      void this.handle(req, res).catch(error => this.error(res, error));
    };
    try {
      for (const host of this.options.tls ? ["127.0.0.1", "::1"] : ["127.0.0.1"]) {
        const server = this.options.tls ? https.createServer({ cert: fs.readFileSync(this.options.tls.cert), key: fs.readFileSync(this.options.tls.key) }, handler) : http.createServer(handler);
        server.headersTimeout = 30_000;
        server.requestTimeout = 60_000;
        try {
          await new Promise<void>((resolve, reject) => { server.once("error", reject); server.listen(this.options.config.port, host, resolve); });
        } catch (error) {
          // IPv6 loopback is optional; IPv4 is the URL the relay and clients use.
          if (host === "::1" && ["EADDRNOTAVAIL", "EAFNOSUPPORT"].includes((error as NodeJS.ErrnoException).code ?? "")) continue;
          this.servers.push(server);
          throw error;
        }
        this.servers.push(server);
        const port = (server.address() as { port: number }).port;
        this.url = `${this.options.tls ? "https" : "http"}://127.0.0.1:${port}`;
      }
      const sibling = fileURLToPath(new URL("./relay.js", import.meta.url));
      const entry = this.options.relayEntry ?? (fs.existsSync(sibling) ? sibling : fileURLToPath(new URL("../../dist/gateway/src/relay.js", import.meta.url)));
      if (!fs.existsSync(entry)) throw new Error("Gateway relay is missing. Build the server before starting it.");
      this.engine = new GatewayEngine(this.options.config, this.registry, { url: this.url, token: this.relayToken, entry, ca: this.options.tls?.cert }, this.options.log, new ConversationStore(this.options.conversationFile));
      await Promise.all([this.models(), ...(this.options.config.warmupOnStart ? [this.warmup()] : [])]);
      this.ready = true;
      return this.url;
    } catch (error) { await this.close(); throw error; }
  }
  private json(res: http.ServerResponse, status: number, data: unknown) {
    res.writeHead(status, { "content-type": "application/json; charset=utf-8" }); res.end(JSON.stringify(data));
  }
  private error(res: http.ServerResponse, error: unknown) {
    if (res.destroyed || res.writableEnded) return;
    const e = error instanceof ApiError ? error : new ApiError(500, "api_error", "Gateway internal error");
    this.options.log?.({ event: "error", status: e.status, type: e.type,
      ...(e.status === 400 || e.status === 409 ? { detail: e.message } : {}) });
    const payload = { type: "error", request_id: res.getHeader("request-id"), error: { type: e.type, message: e.message } };
    if (res.headersSent) { res.write(`event: error\ndata: ${JSON.stringify(payload)}\n\n`); res.end(); }
    else this.json(res, e.status, payload);
  }
  private async body(req: http.IncomingMessage) {
    return new Promise<any>((resolve, reject) => {
      const chunks: Buffer[] = []; let bytes = 0; let failed = false;
      req.on("data", data => {
        if (failed) return; // Drain the socket so the HTTP 413 can reach the client.
        const chunk = Buffer.from(data); bytes += chunk.length;
        if (bytes > this.options.config.maxBodyBytes) {
          failed = true; chunks.length = 0;
          reject(new ApiError(413, "request_too_large", "Request body exceeds configured limit")); return;
        }
        chunks.push(chunk);
      });
      req.on("error", reject);
      req.on("aborted", () => reject(new ApiError(499, "api_error", "Client disconnected during upload")));
      req.on("end", () => {
        if (failed) return;
        try { resolve(JSON.parse(Buffer.concat(chunks).toString("utf8"))); }
        catch { reject(new ApiError(400, "invalid_request_error", "Expected a JSON request body")); }
      });
    });
  }
  private async handle(req: http.IncomingMessage, res: http.ServerResponse) {
    const route = new URL(req.url ?? "/", this.url || "http://localhost").pathname;
    const requestId = `req_agw_${randomUUID()}`;
    res.setHeader("request-id", requestId); res.setHeader("cache-control", "no-store");
    res.on("finish", () => this.options.log?.({ event: "request", method: req.method, path: route, status: res.statusCode, requestId }));
    if (this.stopping) throw new ApiError(503, "api_error", "Gateway shutting down");
    if (route.startsWith("/internal/")) {
      if (!matches(req.headers.authorization, `Bearer ${this.relayToken}`)) throw new ApiError(401, "authentication_error", "Invalid relay token");
      if (req.method !== "POST") throw new ApiError(405, "invalid_request_error", "Expected POST");
      const body = await this.body(req);
      if (typeof body?.session !== "string") throw new ApiError(400, "invalid_request_error", "Missing tool session");
      if (route === "/internal/tools") return this.json(res, 200, this.engine.tools(body.session));
      if (route === "/internal/tool" && typeof body.name === "string" && body.arguments && typeof body.arguments === "object" && !Array.isArray(body.arguments)) {
        return this.json(res, 200, await this.engine.invoke(body.session, body.name, body.arguments));
      }
      throw new ApiError(400, "invalid_request_error", "Invalid relay request");
    }
    const origin = req.headers.origin;
    res.setHeader("vary", "Origin, Access-Control-Request-Headers");
    if (origin && !this.options.config.allowedOrigins.includes(origin)) throw new ApiError(403, "permission_error", "Origin is not allowed");
    if (origin) res.setHeader("access-control-allow-origin", origin);
    res.setHeader("access-control-allow-methods", "GET, POST, OPTIONS");
    res.setHeader("access-control-allow-headers", req.headers["access-control-request-headers"] ?? "x-api-key, authorization, content-type, anthropic-version, anthropic-beta, anthropic-dangerous-direct-browser-access");
    res.setHeader("access-control-allow-private-network", "true");
    res.setHeader("access-control-expose-headers", "request-id, x-agent-gateway-provider, x-agent-gateway-model, x-agent-gateway-usage");
    if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }
    if (req.method === "GET" && ["/", "/health"].includes(route)) return this.json(res, 200, { status: "ok", service: "agent-gateway", mode: "cli", ready: this.ready, sessions: this.engine?.activeSessions ?? 0, instance: process.env.AGENT_GATEWAY_INSTANCE ?? null });
    const valid = matches(typeof req.headers["x-api-key"] === "string" ? req.headers["x-api-key"] : undefined, this.options.token)
      || matches(req.headers.authorization, `Bearer ${this.options.token}`);
    if (!valid) throw new ApiError(401, "authentication_error", "Invalid gateway token");
    if (req.method === "POST" && route === "/gateway/shutdown" && this.options.onShutdown) {
      this.json(res, 200, { status: "stopping" });
      setImmediate(this.options.onShutdown); return;
    }
    if (req.method === "POST" && route === "/gateway/warmup") {
      if (!await this.warmup()) throw new ApiError(502, "api_error", "Provider warmup failed; inspect gateway.log");
      this.json(res, 200, { status: "ready" }); return;
    }
    if (req.method === "GET" && route === "/v1/models") return this.json(res, 200, await this.models());
    if (req.method === "GET" && route === "/gateway/status") return this.json(res, 200, { config: this.options.config, providers: await Promise.all(this.options.config.enabledProviders.map(id => this.registry.require(id).detect())), activeSessions: this.engine.activeSessions });
    if (req.method === "POST" && route === "/v1/messages/count_tokens") throw new ApiError(501, "api_error", "CLI transports do not expose an exact token-count API");
    if (req.method !== "POST" || route !== "/v1/messages") throw new ApiError(404, "not_found_error", "Unknown gateway endpoint");
    const request = parseRequest(await this.body(req));
    const isProbe = request.max_tokens === 1 && !request.stream && !request.system && !request.tools?.length
      && request.messages.length === 1 && request.messages[0]?.role === "user" && request.messages[0]?.content === "x";
    const target = routeModel(request.model, this.options.config, isProbe);
    target.effort = request.output_config?.effort === "max" ? "high" : request.output_config?.effort
      ?? (request.thinking?.type === "disabled" ? "low" : request.thinking?.type === "enabled"
        ? (request.thinking.budget_tokens ?? 8192) > 8192 ? "high" : "medium" : target.effort);
    res.setHeader("x-agent-gateway-provider", target.provider);
    res.setHeader("x-agent-gateway-model", target.model ?? "default");
    res.setHeader("x-agent-gateway-usage", "reported-at-cli-turn-completion");
    this.options.log?.({ event: "route", requestId, provider: target.provider, model: target.model, apiModel: request.model });
    // The official client checks each model using this exact one-token 'x'
    // request with a five-second timeout. Reuse only a recent REAL inference
    // for that exact probe, never fabricated text or cached user conversations.
    const cached = this.probes.get(stable(target));
    if (isProbe && cached && cached.expires > Date.now()) {
      res.setHeader("x-agent-gateway-cache", "verification-probe");
      return this.json(res, 200, { ...cached.message, id: `msg_agw_cached_${randomUUID().replaceAll("-", "")}`, model: request.model,
        usage: { input_tokens: 0, output_tokens: 0 } });
    }
    const owner = createHash("sha256").update(this.options.token).digest("hex");
    if (isProbe) return this.json(res, 200, { ...await this.probe(target, request, owner), model: request.model });
    const running = this.engine.open(request, target, owner);
    res.once("close", () => { if (!res.writableFinished) running.cancel(); });
    if (!request.stream) return this.json(res, 200, await running.segment.result);
    res.writeHead(200, { "content-type": "text/event-stream; charset=utf-8", "x-accel-buffering": "no" });
    const send = (type: string, data: object) => { if (!res.destroyed) res.write(`event: ${type}\ndata: ${JSON.stringify(data)}\n\n`); };
    send("message_start", { type: "message_start", message: { id: `msg_agw_${randomUUID().replaceAll("-", "")}`, type: "message", role: "assistant", model: request.model, content: [], stop_reason: null, stop_sequence: null, usage: { input_tokens: 0, output_tokens: 0 } } });
    let hasText = false;
    const unsubscribe = running.segment.onText(text => {
      if (!hasText) { hasText = true; send("content_block_start", { type: "content_block_start", index: 0, content_block: { type: "text", text: "" } }); }
      send("content_block_delta", { type: "content_block_delta", index: 0, delta: { type: "text_delta", text } });
    });
    const keepalive = setInterval(() => send("ping", { type: "ping" }), 10_000); keepalive.unref();
    try {
      const message = await running.segment.result;
      if (hasText) send("content_block_stop", { type: "content_block_stop", index: 0 });
      let index = hasText ? 1 : 0;
      for (const content of message.content.filter(b => b.type !== "text" || !hasText)) {
        const value = content.type === "tool_use" ? { ...content, input: {} } : content;
        send("content_block_start", { type: "content_block_start", index, content_block: value });
        if (content.type === "tool_use") send("content_block_delta", { type: "content_block_delta", index, delta: { type: "input_json_delta", partial_json: JSON.stringify(content.input) } });
        send("content_block_stop", { type: "content_block_stop", index }); index++;
      }
      send("message_delta", { type: "message_delta", delta: { stop_reason: message.stop_reason, stop_sequence: null }, usage: message.usage });
      send("message_stop", { type: "message_stop" }); res.end();
    } finally { unsubscribe(); clearInterval(keepalive); }
  }
  private async models() {
    if (this.catalog && this.catalog.expires > Date.now()) return this.catalog.value;
    if (!this.catalogLoading) {
      this.catalogLoading = this.computeModels().then(value => {
        this.catalog = { value, expires: Date.now() + 300_000 }; return value;
      }).finally(() => { this.catalogLoading = undefined; });
      void this.catalogLoading.catch(() => {});
    }
    // Native CLI discovery must not block the client's five-second login window.
    return this.catalog?.value ?? this.catalogLoading;
  }
  private async computeModels() {
    const names = new Map<string, string>();
    const config = this.options.config;
    for (const model of config.advertiseCompatibilityModels ? ["claude-opus-4-7", "claude-sonnet-4-6", "claude-haiku-4-5-20251001"] : []) {
      const target = config.aliases[model] ?? config.defaultRoute;
      names.set(model, `${model} → ${target.provider}/${target.model ?? "default"}`);
    }
    for (const [alias, target] of Object.entries(config.aliases)) if (config.enabledProviders.includes(target.provider)) names.set(alias, `${alias} → ${target.provider}/${target.model ?? "default"}`);
    await Promise.all(config.enabledProviders.map(async provider => {
      const adapter = this.registry.require(provider);
      const detection = await adapter.detect();
      if (!detection.installed) return;
      names.set(`${provider}/default`, `${provider} · default CLI model`);
      const catalog = await this.registry.listModels(provider).catch(() => null);
      for (const model of catalog?.models ?? []) names.set(`${provider}/${model.id}`, `${provider} · ${model.name}`);
    }));
    for (const [id, label] of [...names]) if (!/claude|sonnet|opus|haiku|anthropic/i.test(id)) names.set(`${id} [Claude API]`, `${label} · Claude API protocol`);
    const first = `${config.defaultRoute.provider}/default`;
    const rank = (id: string) => id === first ? 0 : id === `${first} [Claude API]` ? 1 : 2;
    const data = [...names].filter(([id]) => !(id.endsWith(" [Claude API]") && config.modelNames[id.slice(0, -" [Claude API]".length)]))
      .sort(([a], [b]) => rank(a) - rank(b) || a.localeCompare(b))
      .map(([id, display_name]) => ({ id: publicModelId(id, config), type: "model", display_name: config.modelNames[id] ?? display_name, created_at: "1970-01-01T00:00:00Z" }));
    return { data, has_more: false, first_id: data[0]?.id ?? null, last_id: data.at(-1)?.id ?? null };
  }
  private cacheProbe(target: Route, message: ApiMessage) {
    if (this.probes.size >= 64) this.probes.delete(this.probes.keys().next().value!);
    this.probes.set(stable(target), { message, expires: Date.now() + this.options.config.probeCacheTtlMs });
  }
  /**
   * Runs one verification probe per route. It is deliberately not tied to the
   * HTTP request: Office abandons the check after five seconds, and a cold CLI
   * start can take longer. Letting it finish fills the cache so Office's retry
   * succeeds, and concurrent probes for the same route share a single turn.
   */
  private probe(target: Route, request: MessagesRequest, owner: string) {
    const key = stable(target);
    let pending = this.probing.get(key);
    if (!pending) {
      pending = this.engine.open(request, target, owner).segment.result.then(message => { this.cacheProbe(target, message); return message; })
        .finally(() => this.probing.delete(key));
      this.probing.set(key, pending);
      void pending.catch(() => {});
    }
    return pending;
  }
  private async warmup() {
    const request = parseRequest({ model: `${this.options.config.defaultRoute.provider}/default`, max_tokens: 1, messages: [{ role: "user", content: "x" }] });
    const target = routeModel(request.model, this.options.config);
    try {
      const owner = createHash("sha256").update(this.options.token).digest("hex");
      await this.probe(target, request, owner);
      this.options.log?.({ event: "warmup_ready", provider: target.provider, model: target.model });
      return true;
    } catch (error) {
      this.options.log?.({ event: "warmup_failed", message: error instanceof Error ? error.message : String(error) });
      return false;
    }
  }
  async close() {
    if (this.stopping) return;
    this.stopping = true;
    await this.engine?.close();
    await this.registry.disposeAll();
    await Promise.all(this.servers.map(server => new Promise<void>(resolve => { server.close(() => resolve()); server.closeAllConnections(); })));
  }
}
