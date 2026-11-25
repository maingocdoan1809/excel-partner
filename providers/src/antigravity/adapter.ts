import { randomUUID } from "node:crypto";
import type { ProviderCapabilities, ProviderConnectionStatus, ProviderEvent, ProviderId } from "../../../protocol/src/index.js";
import { acpModelOption, acpModels, authorizeAcpRequest, contentText, toolNameOf } from "../acp.js";
import { AsyncQueue } from "../async-queue.js";
import { BaseCliAdapter, type CliProbe } from "../base.js";
import { cachedCommands, normalizeCommands } from "../commands.js";
import { discoverAcpCommands } from "../command-discovery.js";
import { effortConfig } from "../effort.js";
import { JsonRpcPeer } from "../jsonrpc.js";
import { CliLoginAttempt, ProviderLoginError, type CliLoginSpec } from "../login.js";
import type { AcpConfigOption } from "../options.js";
import { buildTurnPrompt, SYSTEM_PROMPT } from "../prompt.js";
import { validNativeId, nativeTitle } from "../session-metadata.js";
import type { TerminalAgentSpec } from "../terminal-agent.js";
import { TURN_IDLE_LIMIT_MS } from "../turn-limits.js";
import { MCP_SERVER_NAME, type SendTurnInput, type StartSessionContext } from "../types.js";
import { ensureWorkspaceDir } from "../workspace.js";
import type { ProviderCommand } from "../../../protocol/src/commands.js";

interface AcpSessionResult {
  sessionId?: string;
  configOptions?: AcpConfigOption[];
  models?: { currentModelId?: string; availableModels?: Array<{ modelId: string; name?: string }> };
  modes?: { currentModeId?: string; availableModes?: Array<{ id: string }> };
}
interface ActiveTurn {
  queue: AsyncQueue<ProviderEvent>;
  requestId: string;
  text: string;
  tools: Map<string, string>;
  finishedTools: Set<string>;
  cancel(): void;
}
interface Session {
  peer: JsonRpcPeer;
  nativeId: string;
  config: AcpSessionResult;
  model: string | null;
  systemPrompt: string;
  primed: boolean;
  commands?: ProviderCommand[];
  title?: string;
  active?: ActiveTurn;
  onRuntimeClosed?: () => void;
  detach(): void;
}

/** Official native agy binary, ACP v1 over stdio. */
export class AntigravityAdapter extends BaseCliAdapter {
  readonly id: ProviderId = "antigravity";
  static readonly capabilities: ProviderCapabilities = {
    streaming: true, cancel: true, sessionResume: true, mcp: true, acp: true,
    nativeWebSearch: true, explicitToolNames: false, credits: false,
    loginKind: "browser", logout: false, liveModel: true, liveEffort: true, importSessions: false, fullSchemasOnly: true,
  };
  readonly capabilities = AntigravityAdapter.capabilities;
  static readonly login: CliLoginSpec = {
    name: "Antigravity", args: ["acp", "--login"], pasteBack: true,
    find(output) {
      const uri = /https:\/\/accounts\.google\.com\/[^\s<>"\x1b]+/.exec(output)?.[0];
      return uri ? { uri, code: null } : null;
    },
    instructions: "Sign in with Google, then paste the authorization code from that page here.",
  };
  static readonly terminal: TerminalAgentSpec = {
    command: "agy", label: "Antigravity CLI", client: /antigravity|\bagy\b/i,
    add: server => ["mcp", "add", ...Object.entries(server.env).flatMap(([key, value]) => ["--env", `${key}=${value}`]),
      server.name, "--", server.command, ...server.args],
    remove: name => ["mcp", "remove", name],
  };
  private readonly sessions = new Map<string, Session>();
  private readonly commandCatalog = cachedCommands(() => discoverAcpCommands(this.command, this.args, this.env, ensureWorkspaceDir()));
  private get command() { return this.env.ANTIGRAVITY_BIN ?? "agy"; }
  private get args() { return ["acp", "--disable-telemetry", ...(this.env.ANTIGRAVITY_GEMINI_DIR ? ["--gemini-dir", this.env.ANTIGRAVITY_GEMINI_DIR] : [])]; }
  override startLogin() {
    return new CliLoginAttempt(this.id, this.command, this.env, { ...AntigravityAdapter.login, args: [...this.args, "--login"] });
  }
  protected probe(): CliProbe {
    return { command: this.command, versionArgs: ["--version"],
      installHint: "office-partner providers install antigravity", loginHint: "agy acp --login" };
  }
  private peer(options: Partial<ConstructorParameters<typeof JsonRpcPeer>[0]> = {}) {
    return new JsonRpcPeer({ command: this.command, args: this.args, env: this.env,
      cwd: ensureWorkspaceDir(), versionField: true, ...options });
  }
  private initialize(peer: JsonRpcPeer) {
    return peer.request<{ agentCapabilities?: { sessionCapabilities?: { resume?: object } } }>("initialize", {
      protocolVersion: 1, clientCapabilities: {}, clientInfo: { name: "office-partner", version: "1" },
    }, 15_000);
  }
  private async inspect() {
    const peer = this.peer();
    try {
      await this.initialize(peer);
      return await peer.request<AcpSessionResult>("session/new", { cwd: ensureWorkspaceDir(), mcpServers: [] }, 30_000);
    } finally { await peer.close(); }
  }
  protected override async checkLogin(): Promise<{ status: ProviderConnectionStatus; hint: string | null }> {
    try { await this.inspect(); return { status: "ready", hint: null }; }
    catch (error) {
      if (isAuthError(error)) return { status: "login_required", hint: "Sign in to Antigravity with Google from Settings." };
      throw error;
    }
  }
  async listModels() { return acpModels(await this.inspect()); }
  async listCommands(id?: string) { return (id ? this.sessions.get(id)?.commands : undefined) ?? this.commandCatalog(); }
  private async configure(session: Session, configId: string, value: string) {
    const response = await session.peer.request<AcpSessionResult>("session/set_config_option", { sessionId: session.nativeId, configId, value }, 15_000);
    if (response.configOptions) session.config.configOptions = response.configOptions;
  }
  private async selectModel(session: Session, model: string) {
    const option = acpModelOption(session.config.configOptions);
    if (option) await this.configure(session, option.id, model);
    else await session.peer.request("session/set_model", { sessionId: session.nativeId, modelId: model }, 15_000);
    session.model = model;
  }
  override async startSession(ctx: StartSessionContext): Promise<string> {
    ctx.signal.throwIfAborted();
    if (ctx.resumeSessionId && !validNativeId(ctx.resumeSessionId)) throw new Error("Invalid Antigravity session ID");
    const id = `antigravity-${randomUUID()}`;
    let session: Session | undefined;
    const peer = this.peer({
      onNotification: (method, raw) => { if (session) this.update(session, method, raw); },
      onRequest: (method, params) => authorizeAcpRequest(method, params, ctx.authorizeTool),
    });
    const abort = () => { void this.closeSession(id); void peer.close(); };
    ctx.signal.addEventListener("abort", abort, { once: true });
    try {
      const initialized = await this.initialize(peer);
      const resume = ctx.resumeSessionId;
      const method = resume ? initialized.agentCapabilities?.sessionCapabilities?.resume ? "session/resume" : "session/load" : "session/new";
      const config = await peer.request<AcpSessionResult>(method, {
        ...(resume ? { sessionId: resume } : {}), cwd: ensureWorkspaceDir(),
        mcpServers: [{ type: "stdio", name: MCP_SERVER_NAME, command: ctx.mcp.command, args: ctx.mcp.args,
          env: Object.entries(ctx.mcp.env).map(([name, value]) => ({ name, value })) }],
      }, 30_000);
      const nativeId = resume ?? config.sessionId;
      if (!nativeId || !validNativeId(nativeId)) throw new Error("Antigravity returned no valid session ID");
      session = { peer, nativeId, config, model: acpModels(config).defaultModelId,
        primed: Boolean(resume), systemPrompt: ctx.systemPrompt ?? SYSTEM_PROMPT,
        onRuntimeClosed: ctx.onRuntimeClosed,
        detach: () => ctx.signal.removeEventListener("abort", abort) };
      this.sessions.set(id, session);
      peer.onExit(() => {
        if (this.sessions.get(id) !== session) return;
        this.sessions.delete(id); session!.detach();
        session!.active?.queue.fail(new Error("Antigravity runtime closed unexpectedly"));
        ctx.onRuntimeClosed?.();
      });
      // Do not inherit a saved automatic execution mode from the terminal CLI.
      if (config.modes?.availableModes?.some(mode => mode.id === "default")) {
        await peer.request("session/set_mode", { sessionId: nativeId, modeId: "default" }, 15_000);
      }
      if (ctx.model) await this.selectModel(session, ctx.model);
      ctx.signal.throwIfAborted();
      if (!peer.alive) throw new Error("Antigravity closed during initialization");
      return id;
    } catch (error) {
      this.sessions.delete(id); ctx.signal.removeEventListener("abort", abort); await peer.close();
      if (isAuthError(error)) throw new ProviderLoginError("Antigravity authentication required. Sign in from Settings.");
      throw error;
    }
  }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId), requestId = input.requestId;
    if (!session || !session.peer.alive || session.active) {
      yield { type: "error", requestId, code: "no_session", message: "Antigravity session is unavailable or busy" };
      yield { type: "turn_finished", requestId, reason: "error" }; return;
    }
    const queue = new AsyncQueue<ProviderEvent>();
    const finish = (reason: "completed" | "cancelled" | "error", error?: unknown) => {
      if (queue.isDone) return;
      this.flush(active);
      if (error) queue.push({ type: "error", requestId, code: isAuthError(error) ? "login_required" : "provider_error", message: error instanceof Error ? error.message : String(error) });
      queue.push({ type: "turn_finished", requestId, reason }); queue.end();
    };
    const cancel = () => { session.peer.notify("session/cancel", { sessionId: session.nativeId }); finish("cancelled"); };
    const active: ActiveTurn = { queue, requestId, text: "", tools: new Map(), finishedTools: new Set(), cancel };
    session.active = active;
    input.signal.addEventListener("abort", cancel, { once: true });
    const running = (async () => {
      if (input.signal.aborted) { finish("cancelled"); return; }
      if (input.model && session.model !== input.model) await this.selectModel(session, input.model);
      const effort = effortConfig(session.config.configOptions, input.effort ?? "medium");
      if (effort) await this.configure(session, effort.configId, effort.value);
      if (queue.isDone) return;
      const system = input.systemPrompt ?? session.systemPrompt;
      const user = buildTurnPrompt(input.prompt, input.selection);
      const text = input.command ? `/${input.command.name}${input.command.args ? ` ${input.command.args}` : ""}`
        : session.primed || !system.trim() ? user : `${system}\n\n---\n\n${user}`;
      if (!input.command) session.primed = true;
      const result = await session.peer.request<{ stopReason?: string }>("session/prompt", {
        sessionId: session.nativeId, prompt: [{ type: "text", text }, ...(input.images ?? []).map(image => ({ type: "image", mimeType: image.mediaType, data: image.data }))],
      }, TURN_IDLE_LIMIT_MS, { idle: true });
      const stop = result.stopReason ?? "end_turn";
      if (stop === "cancelled") finish("cancelled");
      else if (stop === "end_turn") finish("completed");
      else finish("error", new Error(`Antigravity stopped the turn (${stop})`));
    })().catch(async error => {
      finish(input.signal.aborted ? "cancelled" : "error", input.signal.aborted ? undefined : error);
      if (/timed out/i.test(String(error))) await this.stopRuntime(input.providerSessionId);
    });
    try {
      yield { type: "session_info", requestId, provider: this.id, providerSessionId: session.nativeId, title: session.title };
      yield { type: "turn_started", requestId };
      for await (const event of queue) yield event;
    }
    finally {
      input.signal.removeEventListener("abort", cancel);
      // A cancelled prompt must settle before another turn uses this runtime.
      if (!queue.isDone) cancel();
      let timer: NodeJS.Timeout | undefined;
      const settled = await Promise.race([running.then(() => true), new Promise<false>(resolve => { timer = setTimeout(() => resolve(false), 5_000); })]);
      clearTimeout(timer);
      if (!settled) await this.stopRuntime(input.providerSessionId);
      if (session.active === active) session.active = undefined;
    }
  }
  override async cancelTurn(id: string) { this.sessions.get(id)?.active?.cancel(); }
  private async stopRuntime(id: string) {
    const session = this.sessions.get(id);
    await this.closeSession(id);
    session?.onRuntimeClosed?.();
  }
  override async closeSession(id: string) {
    const session = this.sessions.get(id);
    if (!session) return;
    this.sessions.delete(id); session.detach(); session.active?.cancel(); await session.peer.close();
  }
  override async dispose() { await Promise.all([...this.sessions.keys()].map(id => this.closeSession(id))); }
  private flush(active: ActiveTurn) {
    if (active.text) active.queue.push({ type: "message_completed", requestId: active.requestId, text: active.text });
    active.text = "";
  }
  private update(session: Session, method: string, raw: unknown) {
    if (method !== "session/update") return;
    const params = raw as { sessionId?: string; update?: Record<string, any> };
    if (params?.sessionId !== session.nativeId || !params.update) return;
    const update = params.update;
    if (update.sessionUpdate === "config_option_update") { session.config.configOptions = update.configOptions; return; }
    if (update.sessionUpdate === "available_commands_update") { session.commands = normalizeCommands(update.availableCommands); return; }
    if (update.sessionUpdate === "session_info_update") { session.title = nativeTitle(update.title) ?? session.title; return; }
    const active = session.active;
    if (!active || active.queue.isDone) return;
    const { queue, requestId } = active;
    if (update.sessionUpdate === "agent_message_chunk" || update.sessionUpdate === "agent_thought_chunk") {
      const text = update.content?.type === "text" ? String(update.content.text ?? "") : "";
      if (!text) return;
      const thinking = update.sessionUpdate === "agent_thought_chunk";
      if (!thinking) active.text += text;
      queue.push({ type: thinking ? "thinking_delta" : "message_delta", requestId, text });
    } else if (update.sessionUpdate === "tool_call" || update.sessionUpdate === "tool_call_update") {
      if (typeof update.toolCallId !== "string") return;
      this.flush(active);
      const callId = update.toolCallId;
      if (!active.tools.has(callId)) {
        active.tools.set(callId, toolNameOf(update));
        queue.push({ type: "tool_started", requestId, callId, tool: active.tools.get(callId)!, args: update.rawInput ?? {} });
      }
      if (["completed", "failed"].includes(update.status) && !active.finishedTools.has(callId)) {
        active.finishedTools.add(callId);
        queue.push({ type: "tool_finished", requestId, callId, tool: active.tools.get(callId)!, ok: update.status === "completed",
          error: update.status === "failed" ? contentText(update).slice(0, 500) || "Tool call failed" : null });
      }
    }
  }
}
function isAuthError(error: unknown) { return /authentication required|not logged in|unauthenticated|credentials.*(?:missing|expired)|\b401\b/i.test(String(error)); }
