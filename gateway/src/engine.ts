import { clientToolResult, CONVERSION_ERROR } from "./tool-result.js";
import { gatewayToolInstructions } from "./tool-instructions.js";
import { matchesHistory, snapshotHistory, type HistorySnapshot } from "./history.js";
import { conversationKey, ConversationStore, type ConversationRecord } from "./conversations.js";
import { randomUUID, createHash } from "node:crypto";
import { CLIENT_TOOL_CONTRACT_REVISION } from "../../providers/src/client-tools.js";
import type { ProviderRegistry } from "../../providers/src/registry.js";
import { ProviderSessionManager, type ProviderSessionState } from "../../providers/src/session-manager.js";
import type { TokenUsage } from "../../protocol/src/index.js";
import { ApiError, stable, transcriptPrompt, turnImages, type ApiMessage, type ChatMessage, type ClientTool, type ContentBlock, type MessagesRequest } from "./protocol.js";
import type { GatewayConfig, Route } from "./config.js";

function fingerprint(value: unknown) { return createHash("sha256").update(stable(value)).digest("hex"); }
interface Pending {
  id: string; name: string; input: Record<string, unknown>; exposed: boolean;
  resolve(value: unknown): void; reject(error: Error): void;
}
export class Segment {
  text = "";
  settled = false;
  private listeners = new Set<(text: string) => void>();
  private resolve!: (value: ApiMessage) => void;
  private reject!: (error: Error) => void;
  readonly result = new Promise<ApiMessage>((resolve, reject) => { this.resolve = resolve; this.reject = reject; });
  constructor(readonly model: string) { void this.result.catch(() => {}); }
  onText(listener: (text: string) => void): () => void {
    if (this.text) listener(this.text);
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
  append(text: string) {
    if (this.settled || !text) return;
    this.text += text;
    for (const listener of this.listeners) listener(text);
  }
  finish(reason: ApiMessage["stop_reason"], tools: ContentBlock[], usage?: TokenUsage) {
    if (this.settled) return;
    this.settled = true;
    const content: ContentBlock[] = [...(this.text ? [{ type: "text", text: this.text }] : []), ...tools];
    const message: ApiMessage = { id: `msg_agw_${randomUUID().replaceAll("-", "")}`, type: "message", role: "assistant", model: this.model,
      content: content.length ? content : [{ type: "text", text: "" }], stop_reason: reason, stop_sequence: null,
      usage: { input_tokens: usage?.uncachedInputTokens ?? (usage ? Math.max(0, usage.inputTokens - usage.cachedInputTokens - (usage.cacheWriteTokens ?? 0)) : 0),
        output_tokens: usage?.outputTokens ?? 0, ...(usage ? { cache_read_input_tokens: usage.cachedInputTokens, cache_creation_input_tokens: usage.cacheWriteTokens ?? 0 } : {}) } };
    this.resolve(message);
    this.listeners.clear();
    return message;
  }
  fail(error: Error) { if (this.settled) return; this.settled = true; this.reject(error); this.listeners.clear(); }
}

interface Job {
  conversationKey?: string;
  id: string; owner: string; route: Route; request: MessagesRequest; configHash: string; expected?: HistorySnapshot;
  state: ProviderSessionState; controller: AbortController; turnController: AbortController; pending: Map<string, Pending>;
  idle: boolean; completedHistory?: HistorySnapshot; historyLength?: number; idleSince?: number;
  exposed: Set<string>; segment: Segment; usage?: TokenUsage; timer?: NodeJS.Timeout; batch?: NodeJS.Timeout;
  timeout?: NodeJS.Timeout; disposed: boolean;
  /** True once this turn called any tool; text-only turns can be retried safely. */
  sideEffects: boolean;
  traceId: string; startedAt: number; segmentStartedAt: number; firstText: boolean; lastTextAt?: number; toolWaitAt?: number;
}

/** Maps API history to conversations; provider lifecycle is shared with the add-in. */
export class GatewayEngine {
  private readonly jobs = new Map<string, Job>();
  // Closed idle runtimes retain only bounded native identity + history digests.
  private readonly parked = new Map<string, { job: Job; expires: number }>();
  private readonly calls = new Map<string, Job>();
  private readonly running = new Set<Promise<void>>();
  private readonly retiring = new Set<Promise<void>>();
  private closed = false;
  private readonly sessions: ProviderSessionManager;
  constructor(readonly config: GatewayConfig, private readonly registry: ProviderRegistry,
    private readonly relay: { url: string; token: string; entry: string; ca?: string },
    private readonly log: (entry: object) => void = () => {},
    private readonly conversations = new ConversationStore()) {
    this.sessions = new ProviderSessionManager(registry, async () => {}, {
      debug: (action, details) => log({ event: "provider_session", action, ...details }),
      warn: (action, details) => log({ event: "provider_session_warning", action, ...details }),
    });
  }
  get activeSessions() { return this.jobs.size; }
  tools(id: string) {
    const job = this.jobs.get(id);
    if (!job || job.disposed) throw new ApiError(404, "not_found_error", "Tool session expired");
    return { tools: this.clientTools(job.request).map(t => ({ name: t.name, description: t.description ?? "", inputSchema: t.input_schema })) };
  }
  private clientTools(request: MessagesRequest): ClientTool[] {
    if (request.tool_choice?.type === "none") return [];
    if (request.tool_choice?.type === "tool") return (request.tools ?? []).filter(t => t.name === request.tool_choice?.name);
    return request.tools ?? [];
  }
  private configHash(request: MessagesRequest, route: Route) {
    // Preserve semantic changes, while ignoring representation and list ordering
    // that do not alter the tools/system text delivered to the provider.
    const system = typeof request.system === "string" ? request.system
      : (request.system ?? []).filter(b => b.type === "text").map(b => String(b.text ?? "")).join("\n");
    const choice = { type: request.tool_choice?.type ?? "auto", disable_parallel_tool_use: request.tool_choice?.disable_parallel_tool_use ?? false,
      ...(request.tool_choice?.type === "tool" ? { name: request.tool_choice.name } : {}) };
    return fingerprint({ contract: CLIENT_TOOL_CONTRACT_REVISION, route, tools: [...this.clientTools(request)].sort((a, b) => a.name.localeCompare(b.name)), tool_choice: choice, system });
  }
  private turnContent(request: MessagesRequest, route: Route, messages: ChatMessage[]) {
    const clientSystem = typeof request.system === "string" ? request.system : (request.system ?? []).filter(b => b.type === "text").map(b => String(b.text ?? "")).join("\n");
    const officeJs = this.clientTools(request).some(tool => tool.name === "execute_office_js");
    const system = [clientSystem, `<gateway_tool_contract>\n${gatewayToolInstructions(route.provider, officeJs)}\n</gateway_tool_contract>`].filter(Boolean).join("\n\n");
    const prompt = transcriptPrompt({ ...request, messages });
    if (system.length + prompt.length > this.config.maxPromptChars) throw new ApiError(413, "invalid_request_error", "Conversation text exceeds the gateway prompt limit. Compact the conversation or start a new chat; no provider turn was sent.");
    return { system, prompt };
  }

  open(request: MessagesRequest, route: Route, owner: string): { segment: Segment; cancel: () => void } {
    if (this.closed) throw new ApiError(503, "api_error", "Gateway shutting down");
    const key = conversationKey(request, owner, route.provider);
    const last = request.messages.at(-1)!;
    const results = last.role === "user" && Array.isArray(last.content) ? last.content.filter(b => b.type === "tool_result") : [];
    // Office can seed a new request with tools it already executed locally
    // (workbook context/bootstrap). Those results are ordinary supplied history,
    // not continuations of a CLI turn owned by this gateway.
    const hasGatewayResults = results.some(r => typeof r.tool_use_id === "string" && r.tool_use_id.startsWith("toolu_agw_"));
    if (hasGatewayResults) {
      const ids = results.map(r => String(r.tool_use_id ?? ""));
      const job = this.calls.get(ids[0]!);
      if (!job || job.disposed) throw new ApiError(409, "invalid_request_error", "Tool continuation expired or is unknown. Start a new conversation; this request was not replayed.");
      if (job.owner !== owner || job.conversationKey !== key || job.configHash !== this.configHash(request, route)) throw new ApiError(409, "invalid_request_error", "Tool continuation belongs to a different conversation, token, model, system prompt or tool configuration");
      if (new Set(ids).size !== ids.length || ids.length !== job.exposed.size || ids.some(id => !job.exposed.has(id) || this.calls.get(id) !== job)) {
        throw new ApiError(400, "invalid_request_error", "Supply exactly one result for every tool_use in the preceding response");
      }
      const incomingHistory = snapshotHistory(request.messages.slice(0, -1));
      const exactHistory = job.expected && matchesHistory(job.expected, incomingHistory);
      // Office context_snip modifies prior messages. Its live owned call ID,
      // explicit chat identity and exact last assistant response prove this
      // continuation without guessing a branch from a partial transcript.
      const ownedSnip = key && ids.some(id => job.pending.get(id)?.name === "context_snip")
        && job.expected && matchesHistory(job.expected.slice(-1), incomingHistory.slice(-1));
      if (!exactHistory && !ownedSnip) throw new ApiError(409, "invalid_request_error", "Conversation history does not match the pending tool response");
      const converted = results.map(result => clientToolResult(result.content, result.is_error === true));
      clearTimeout(job.timer);
      job.request = request;
      job.segment = new Segment(request.model);
      job.exposed.clear();
      job.segmentStartedAt = performance.now();
      job.firstText = false;
      job.lastTextAt = undefined;
      this.trace(job, "tool_result", { count: results.length, errorCount: converted.filter(result => result.isError).length,
        historyRebased: !exactHistory,
        clientErrorCount: results.filter(result => result.is_error === true).length,
        conversionErrorCount: converted.filter(result => result._meta?.[CONVERSION_ERROR]).length,
        clientWaitMs: job.toolWaitAt === undefined ? 0 : Math.round(performance.now() - job.toolWaitAt) });
      this.armRequest(job);
      // Install the new segment before resolving tool handlers: the CLI may resume immediately.
      for (const [index, result] of results.entries()) {
        const id = String(result.tool_use_id);
        const pending = job.pending.get(id)!;
        job.pending.delete(id); this.calls.delete(id);
        pending.resolve(converted[index]!);
      }
      if ([...job.pending.values()].some(p => !p.exposed)) this.scheduleTools(job);
      return { segment: job.segment, cancel: this.cancellation(job) };
    }
    const configHash = this.configHash(request, route);
    // A full completed prefix proves identity without requiring a client-specific
    // conversation header. Never guess from just token, model or last user text.
    const prefixes = new Map<number, HistorySnapshot>();
    const prefixHash = (length: number) => {
      if (!prefixes.has(length)) prefixes.set(length, snapshotHistory(request.messages.slice(0, length)));
      return prefixes.get(length)!;
    };
    for (const [id, record] of this.parked) if (record.expires <= Date.now()) this.parked.delete(id);
    const retained = [...this.jobs.values(), ...[...this.parked.values()].map(record => record.job)];
    const identified = key ? retained.find(job => job.conversationKey === key) : undefined;
    const saved = key ? this.conversations.get(key) : undefined;
    if (identified && !identified.idle) throw new ApiError(409, "invalid_request_error", "This conversation already has an active turn or unresolved client tools");
    if (key && !identified && saved && saved.status !== "idle") throw new ApiError(409, "invalid_request_error", "This conversation has an interrupted or failed turn. Its actions may already have executed; automatic replay is disabled. Start a new chat after inspecting workbook state.");
    const candidates = retained.filter(job => job.idle && job.owner === owner
      && job.route.provider === route.provider && job.historyLength !== undefined
      && (key ? job === identified : job.conversationKey === undefined)
      && request.messages.length > job.historyLength
      && request.messages[job.historyLength]?.role === "user"
      && job.completedHistory && matchesHistory(job.completedHistory, prefixHash(job.historyLength)));
    if (identified && !candidates.length) throw new ApiError(409, "invalid_request_error", "Conversation ID matches, but supplied history is edited, truncated, or repeats a completed turn. No provider turn was sent.");
    if (candidates.length === 1) {
      const job = candidates[0]!;
      const appended = request.messages.slice(job.historyLength!);
      this.turnContent(request, route, appended);
      if (this.parked.has(job.id)) {
        this.makeRoom();
        this.parked.delete(job.id);
        job.disposed = false;
        job.controller = new AbortController();
        this.jobs.set(job.id, job);
      }
      clearTimeout(job.timer);
      job.idle = false;
      job.request = request;
      job.route = route;
      job.configHash = configHash;
      job.segment = new Segment(request.model);
      job.turnController = new AbortController();
      job.usage = undefined;
      job.expected = undefined;
      this.startTurn(job, appended, key ? "conversation_id" : "history_match");
      return { segment: job.segment, cancel: this.cancellation(job) };
    }
    // Idle workers are a bounded cache; active/tool-waiting turns are never evicted.
    let checkpoint: ConversationRecord | undefined;
    if (saved) {
      const length = saved.history?.length;
      if (!saved.nativeId || length === undefined || request.messages.length <= length || request.messages[length]?.role !== "user"
        || !matchesHistory(saved.history!, prefixHash(length))) throw new ApiError(409, "invalid_request_error", "Stored conversation checkpoint does not match supplied history or has no confirmed native ID. No provider turn was sent.");
      checkpoint = saved;
    }
    this.turnContent(request, route, checkpoint ? request.messages.slice(checkpoint.history!.length) : request.messages);
    this.makeRoom();
    this.registry.require(route.provider);
    const id = checkpoint?.id ?? randomUUID();
    const job: Job = { id, owner, route, request, configHash, conversationKey: key,
      state: { id, provider: route.provider, providerSessionId: null, providerNativeSessionId: checkpoint?.nativeId ?? null,
        providerChatSessionId: null, providerModel: null, providerEffort: route.effort,
        systemPromptRevision: null, capabilities: { webSearch: false, playwright: false } },
      controller: new AbortController(), turnController: new AbortController(), pending: new Map(), exposed: new Set(),
      segment: new Segment(request.model), disposed: false, idle: false, sideEffects: false,
      traceId: "", startedAt: 0, segmentStartedAt: 0, firstText: false };
    this.jobs.set(job.id, job);
    job.completedHistory = checkpoint?.history;
    job.historyLength = checkpoint?.history?.length;
    this.startTurn(job, checkpoint ? request.messages.slice(job.historyLength!) : request.messages,
      checkpoint ? "persisted_conversation_id" : key ? "new_conversation_id" : candidates.length > 1 ? "ambiguous_history" : "new_history");
    return { segment: job.segment, cancel: this.cancellation(job) };
  }

  private makeRoom() {
    if (this.jobs.size >= this.config.maxSessions) {
      const oldest = [...this.jobs.values()].filter(job => job.idle).sort((a, b) => a.idleSince! - b.idleSince!)[0];
      if (oldest) {
        const cleanup = this.dispose(oldest, undefined, "capacity_eviction");
        this.retiring.add(cleanup);
        void cleanup.then(() => this.retiring.delete(cleanup), () => this.retiring.delete(cleanup));
      }
      else throw new ApiError(429, "rate_limit_error", "Gateway concurrency limit reached");
    }
  }

  private cancellation(job: Job) {
    const segment = job.segment;
    return () => {
      if (!job.idle && !job.disposed && job.segment === segment) void this.dispose(job, new ApiError(499, "api_error", "Client disconnected"), "client_disconnect");
    };
  }
  private trace(job: Job, phase: string, fields: Record<string, unknown> = {}) {
    this.log({ event: "gateway_turn", sessionId: job.id, traceId: job.traceId, provider: job.route.provider, phase,
      elapsedMs: Math.round(performance.now() - job.startedAt), ...fields });
  }
  private startTurn(job: Job, messages: ChatMessage[], match: string) {
    job.traceId = randomUUID();
    job.sideEffects = false;
    job.startedAt = job.segmentStartedAt = performance.now();
    job.firstText = false;
    job.lastTextAt = undefined;
    this.trace(job, "start", { match, identitySource: job.conversationKey ? "metadata.user_id.session_id" : "history", historyMessages: job.request.messages.length, appendedMessages: messages.length, tools: this.clientTools(job.request).length });
    this.armRequest(job);
    const task = Promise.resolve().then(() => { this.checkpoint(job, "inflight"); return this.run(job, messages); })
      .catch(error => this.dispose(job, error instanceof ApiError ? error : new ApiError(502, "api_error", error instanceof Error ? error.message : String(error))));
    this.running.add(task);
    void task.then(() => this.running.delete(task), () => this.running.delete(task));
  }

  private checkpoint(job: Job, status: ConversationRecord["status"]) {
    if (!job.conversationKey) return;
    this.conversations.set({ key: job.conversationKey, id: job.id,
      nativeId: job.state.providerNativeSessionId, history: job.completedHistory,
      status, updated: Date.now() });
  }

  /** A turn that never called a tool cannot have changed the client, so Stop or a
   * dropped connection returns the chat to its last completed state for a retry. */
  private restoreCheckpoint(job: Job) {
    if (!job.conversationKey) return;
    if (job.completedHistory && job.state.providerNativeSessionId) this.checkpoint(job, "idle");
    else this.conversations.delete(job.conversationKey);
  }

  private armRequest(job: Job) {
    clearTimeout(job.timeout);
    job.timeout = setTimeout(() => { void this.dispose(job, new ApiError(504, "api_error", "CLI request timed out"), "request_timeout"); }, this.config.requestTimeoutMs);
    job.timeout.unref();
  }
  private scheduleTools(job: Job) {
    if (job.batch || job.segment.settled) return;
    job.batch = setTimeout(() => {
      job.batch = undefined;
      if (job.disposed || job.segment.settled) return;
      const all = [...job.pending.values()].filter(p => !p.exposed);
      const pending = job.request.tool_choice?.disable_parallel_tool_use ? all.slice(0, 1) : all;
      if (!pending.length) return;
      const blocks = pending.map(p => { p.exposed = true; job.exposed.add(p.id); return { type: "tool_use", id: p.id, name: p.name, input: p.input }; });
      const content: ContentBlock[] = [...(job.segment.text ? [{ type: "text", text: job.segment.text }] : []), ...blocks];
      job.expected = snapshotHistory([...job.request.messages, { role: "assistant", content }]);
      clearTimeout(job.timeout);
      job.timer = setTimeout(() => { void this.dispose(job, new ApiError(408, "api_error", "Timed out waiting for client tool results"), "tool_timeout"); }, this.config.continuationTimeoutMs);
      job.timer.unref();
      job.toolWaitAt = performance.now();
      this.trace(job, "tool_wait", { count: blocks.length, tools: pending.map(p => p.name), segmentMs: Math.round(performance.now() - job.segmentStartedAt) });
      job.segment.finish("tool_use", blocks);
    }, 15);
  }
  async invoke(id: string, name: string, input: Record<string, unknown>): Promise<unknown> {
    const job = this.jobs.get(id);
    if (!job || job.disposed || job.idle) throw new ApiError(404, "not_found_error", "Tool session expired");
    if (!this.clientTools(job.request).some(t => t.name === name)) throw new ApiError(403, "permission_error", "Client did not register this tool");
    job.sideEffects = true;
    const callId = `toolu_agw_${randomUUID().replaceAll("-", "")}`;
    const promise = new Promise((resolve, reject) => {
      job.pending.set(callId, { id: callId, name, input, exposed: false, resolve, reject });
      this.calls.set(callId, job);
    });
    this.scheduleTools(job);
    return promise;
  }

  private async run(job: Job, messages: ChatMessage[]) {
    const names = this.clientTools(job.request).map(t => t.name);
    let { system, prompt } = this.turnContent(job.request, job.route, !job.state.providerNativeSessionId && !job.state.providerSessionId ? job.request.messages : messages);
    let opened = false;
    const nativeSessionId = job.state.providerNativeSessionId;
    await this.sessions.ensure(job.state, { model: job.route.model, effort: job.route.effort,
      chatSessionId: job.state.providerChatSessionId ?? job.id, capabilities: { webSearch: false, playwright: false }, systemPromptRevision: job.configHash }, () => { opened = true; return { workbookConnectionId: job.id, sessionId: job.id, model: job.route.model, effort: job.route.effort,
      systemPrompt: system, systemPromptMode: "replace", signal: job.controller.signal,
      onDiagnostic: event => this.trace(job, "provider_status", { statusPhase: event.phase, attempt: event.attempt, status: event.status, errorType: event.errorType, nextRetryAt: event.nextRetryAt }),
      requestUserInput: async () => ({ cancelled: true, answers: {} }), authorizeTool: async () => false,
      mcp: { bridgeUrl: this.relay.url, sessionId: job.id, token: this.relay.token, command: process.execPath, args: [this.relay.entry], toolNames: names,
        env: { AGENT_GATEWAY_RELAY_URL: this.relay.url, AGENT_GATEWAY_RELAY_TOKEN: this.relay.token, AGENT_GATEWAY_RELAY_SESSION: job.id, AGENT_GATEWAY_PROVIDER: job.route.provider,
          ...(names.includes("execute_office_js") ? { AGENT_GATEWAY_OFFICE_JS: "1" } : {}),
          ...(this.relay.ca ? { NODE_EXTRA_CA_CERTS: this.relay.ca } : {}) } } }; });
    if (job.disposed) return;
    if (opened && !nativeSessionId) ({ system, prompt } = this.turnContent(job.request, job.route, job.request.messages));
    this.trace(job, "runtime_ready", { runtime: opened ? nativeSessionId ? "resumed" : "new" : "reused" });
    let streamed = "";
    let reason = "error";
    for await (const event of this.sessions.sendTurn(job.state, { requestId: randomUUID(), model: job.route.model, effort: job.route.effort,
      systemPrompt: system, prompt, images: turnImages(messages), selection: null, signal: job.turnController.signal })) {
      if (job.disposed) break;
      if ((event.type === "message_delta" || event.type === "message_completed") && event.text) {
        job.lastTextAt = performance.now();
        if (!job.firstText) {
          job.firstText = true;
          this.trace(job, "first_text", { segmentMs: Math.round(performance.now() - job.segmentStartedAt) });
        }
      }
      if (event.type === "message_delta") { streamed += event.text; job.segment.append(event.text); }
      if (event.type === "message_completed") {
        if (!streamed && event.text) job.segment.append(event.text);
        streamed = "";
      }
      if (event.type === "tool_started") job.sideEffects = true;
      if (event.type === "tool_started") this.trace(job, "provider_tool_started", { tool: event.tool });
      if (event.type === "tool_finished") this.trace(job, "provider_tool_finished", { tool: event.tool, ok: event.ok });
      if (event.type === "usage_update") job.usage = event.usage;
      if (event.type === "error") throw new ApiError(event.code === "login_required" ? 401 : 502, event.code === "login_required" ? "authentication_error" : "api_error", event.message);
      if (event.type === "turn_finished") reason = event.reason;
    }
    if (job.disposed) return;
    if (reason !== "completed") throw new ApiError(502, "api_error", `Provider turn ended: ${reason}`);
    if (job.pending.size) throw new ApiError(502, "api_error", "Provider finished before pending client tools were resolved");
    clearTimeout(job.timeout);
    this.trace(job, "completed", { segmentMs: Math.round(performance.now() - job.segmentStartedAt),
      afterTextMs: job.lastTextAt === undefined ? null : Math.round(performance.now() - job.lastTextAt) });
    const content = [{ type: "text", text: job.segment.text }];
    job.completedHistory = snapshotHistory([...job.request.messages, { role: "assistant", content }]);
    job.historyLength = job.request.messages.length + 1;
    this.checkpoint(job, "idle");
    job.segment.finish("end_turn", [], job.usage);
    job.idle = true;
    job.idleSince = Date.now();
    job.timer = setTimeout(() => { void this.dispose(job); }, this.config.idleTimeoutMs);
    job.timer.unref();
  }

  private async dispose(job: Job, error?: Error, reason = error ? "gateway_error" : "idle_expired") {
    if (job.disposed) return;
    job.disposed = true;
    if (!job.idle && job.conversationKey) {
      try { if (job.sideEffects) this.checkpoint(job, "blocked"); else this.restoreCheckpoint(job); }
      catch { this.log({ event: "conversation_checkpoint_error", action: "block_failed" }); }
    }
    clearTimeout(job.timer); clearTimeout(job.timeout); clearTimeout(job.batch);
    if (error) job.segment.fail(error);
    this.jobs.delete(job.id);
    for (const pending of job.pending.values()) { this.calls.delete(pending.id); pending.reject(error ?? new Error("Tool session ended")); }
    job.pending.clear();
    job.turnController.abort();
    job.controller.abort();
    if (!job.idle) {
      try { await this.sessions.cancel(job.state); } catch { /* transport already stopped */ }
    }
    const preserve = !this.closed && job.idle && Boolean(job.state.providerNativeSessionId)
      && (reason === "idle_expired" || reason === "capacity_eviction");
    // Capture state independently: close() mutates it before awaiting adapter cleanup.
    const parkedState = preserve ? { ...job.state, providerSessionId: null, systemPromptRevision: null } : undefined;
    const cleanup = this.sessions.close(job.state, reason, preserve);
    if (parkedState) {
      const parkedJob = { ...job, state: parkedState, request: { ...job.request, messages: [] }, segment: new Segment(job.request.model) };
      this.parked.set(job.id, { job: parkedJob, expires: Date.now() + 24 * 60 * 60_000 });
      while (this.parked.size > this.config.maxSessions * 4) this.parked.delete(this.parked.keys().next().value!);
    }
    await cleanup;
  }
  async close() {
    this.closed = true;
    this.parked.clear();
    await Promise.all([...this.jobs.values()].map(job => this.dispose(job, new ApiError(503, "api_error", "Gateway shutting down"), "shutdown")));
    // Initializers cancelled during startup still own a process until ensure()
    // discards their late result. Wait for that shared-manager cleanup as well.
    await Promise.all([...this.running, ...this.retiring]);
  }
}
