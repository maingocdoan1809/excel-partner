import { TURN_IDLE_LIMIT_MS } from "../turn-limits.js";
import { randomUUID } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderEvent,
  ProviderId,
  TokenUsage,
} from "../../../protocol/src/index.js";
import { acpPermissionOutcome, contentText, toolNameOf } from "../acp.js";
import { AsyncQueue } from "../async-queue.js";
import { BaseCliAdapter, cliEnvironment, execCli, type CliProbe } from "../base.js";
import { ProviderLoginError } from "../login.js";
import { OpenCodeAuthServer, OpenCodeLoginAttempt, type OpenCodeAuthApi } from "./auth.js";
import { effortConfig, type EffortConfigOption } from "../effort.js";
import { acpOptions, type AcpConfigOption } from "../options.js";
import type { ProviderOption } from "../../../protocol/src/provider-options.js";
import { JsonRpcPeer } from "../jsonrpc.js";
import { buildTurnPrompt, SYSTEM_PROMPT } from "../prompt.js";
import { CumulativeUsage } from "../usage.js";
import { MCP_SERVER_NAME, type ProviderLoginAttempt, type ProviderModelCatalog, type SendTurnInput, type StartSessionContext } from "../types.js";
import type { LoginAccount, LoginChoice } from "../../../protocol/src/provider-login.js";
import { ensureWorkspaceDir } from "../workspace.js";
import { nativeTitle, validNativeId } from "../session-metadata.js";
import { parseOpencodeExport } from "./usage.js";
import { reserveOpencodePort, watchOpencodeQuestions } from "./questions.js";
import { watchOpencodeStatus, type OpencodeStatus } from "./status.js";
import { cachedCommands, normalizeCommands } from "../commands.js";
import { discoverAcpCommands } from "../command-discovery.js";
import type { ProviderCommand } from "../../../protocol/src/commands.js";

/** One entry of the CLI's `session/new` / `config_option_update` config list. */
export interface OpencodeConfigOption {
  id?: string;
  name?: string;
  category?: string;
  type?: string;
  currentValue?: string;
  options?: Array<{ value?: string; name?: string }>;
}

interface OpencodeSession {
  commands?: ProviderCommand[];
  questionApi?: { port: number; password: string; cwd: string; ask: StartSessionContext["requestUserInput"] };
  peer: JsonRpcPeer;
  /** ACP session id returned by `session/new`. */
  acpSessionId: string;
  /** Config options the CLI offers (model, effort, mode). */
  configOptions: OpencodeConfigOption[];
  /** Model currently selected on the CLI side. */
  model: string | null;
  /** True once the system prompt has been delivered with the first prompt. */
  primed: boolean;
  active: ActiveTurn | null;
  usage?: CumulativeUsage;
  usageBaselinePending?: boolean;
  title?: string;
  /** Temp directory holding the permission config handed to the CLI. */
  configDir: string;
  v2?: boolean;
  /**
   * Usage is read from the CLI after every turn. Only an API client needs it that often (it answers with the figure);
   * for a task pane it is read when someone opens the usage (`sessionUsage`), since reading it runs the CLI.
   */
  reportsUsage: boolean;
  diagnostic?: StartSessionContext["onDiagnostic"];
  /** The `session/prompt` still waiting for its answer; after a cancel it settles when the CLI confirms it stopped. */
  prompting?: Promise<unknown>;
}

/** How long a cancelled prompt may take to be confirmed before the CLI is considered stuck and is closed. */
const CANCEL_GRACE_MS = 5_000;

/** How long OpenCode may retry a failing model call, and how long a healthy model may stay silent, before the first output. */
const FIRST_RESPONSE_RETRY_LIMIT_MS = 120_000;
const FIRST_RESPONSE_LIMIT_MS = 600_000;

interface ActiveTurn {
  queue: AsyncQueue<ProviderEvent>;
  requestId: string;
  /** Assistant text accumulated since the last tool call / message boundary. */
  text: string;
  tools: Map<string, string>;
  finishedTools: Set<string>;
  receivedOutput?: boolean;
  /** OpenCode reports that the model call failed and is being retried. */
  retrying?: boolean;
  stopWaiting?: () => void;
}

/**
 * Turn the CLI's model config option into the picker catalog.
 * OpenCode reports enabled models as `provider/model` (a `/variant` suffix
 * also pins the effort).
 */
export function parseOpencodeModels(options: OpencodeConfigOption[] | undefined): ProviderModelCatalog {
  const option = options?.find((item) => item.category === "model" || item.id === "model");
  const seen = new Set<string>();
  const models = (option?.options ?? []).flatMap((entry) => {
    const id = entry.value;
    if (!id || seen.has(id)) return [];
    seen.add(id);
    return [{ id, name: entry.name || id }];
  });
  return { models, defaultModelId: option?.currentValue ?? null };
}

/**
 * OpenCode CLI via ACP (`opencode acp`, JSON-RPC over stdio).
 *
 * Verified against opencode 1.19.3:
 *   initialize -> session/new (mcpServers: [excel]) -> session/prompt
 *   notifications: session/update (agent_message_chunk, tool_call,
 *   tool_call_update, config_option_update); cancel = session/cancel.
 *
 * OpenCode runs its own tools (shell, file edits) without asking first, so
 * the CLI is started with a permission config that turns every tool into a
 * question: only the Excel MCP tools are pre-approved, and this adapter
 * forwards other permission requests to the pane’s Ask/Auto approval flow.
 */
/** Major version >= 2. Anchored so "1.2.3" is not mistaken for "2.". */
export function isOpencodeV2(version: string): boolean {
  return /^\D*(?:[2-9]|\d{2,})\./.test(version);
}

export class OpencodeAdapter extends BaseCliAdapter {
  private readonly commandCatalog = cachedCommands(() => discoverAcpCommands(this.command, ["acp"], this.env, ensureWorkspaceDir()));
  async listCommands(id?: string) {
    const commands = await ((id ? this.sessions.get(id)?.commands : undefined) ?? this.commandCatalog());
    const version = (await this.detect()).version ?? "";
    return isOpencodeV2(version) && !commands.some(command => command.name === "skills")
      ? [...commands, ...normalizeCommands([{ name: "skills", description: "List available OpenCode skills" }])] : commands;
  }
  private titleCache?: { at: number; value: Promise<Map<string, string>> };
  /** The CLI's own export of one conversation. */
  async sessionUsage(id: string): Promise<{ usage: TokenUsage; title?: string }> {
    if (!validNativeId(id)) throw new Error("Invalid OpenCode session ID");
    return this.originalUsage(id);
  }

  async sessionTitle(id: string, options?: { cheapOnly?: boolean }): Promise<string | undefined> {
    // Reading titles runs the CLI; a caller that is not showing them to anyone right now does not ask.
    if (options?.cheapOnly || !validNativeId(id)) return undefined;
    if (!this.titleCache || Date.now() - this.titleCache.at > 5000) {
      this.titleCache = { at: Date.now(), value: execCli(this.command,
        ["session", "list", "--format", "json", "--max-count", "100", "--pure"], this.env, 5000).then(result => {
        const titles = new Map<string, string>();
        try {
          if (result.code !== 0) return titles;
          for (const item of JSON.parse(result.stdout)) {
            const title = nativeTitle(item.title);
            if (typeof item.id === "string" && title) titles.set(item.id, title);
          }
        } catch { /* Session titles remain optional on older CLI versions. */ }
        return titles;
      }) };
    }
    return (await this.titleCache.value).get(id);
  }

  private async originalUsage(id: string, v2 = false): Promise<ReturnType<typeof parseOpencodeExport>> {
    const result = await execCli(this.command, v2 ? ["session", "export", id] : ["export", id, "--pure"], this.env, 10_000);
    if (result.code !== 0) throw new Error("OpenCode original session export is unavailable");
    return parseOpencodeExport(JSON.parse(result.stdout), id, v2);
  }
  readonly id: ProviderId = "opencode";
  /** How OpenCode reaches the client's tools: through its Code Mode sandbox rather than as ordinary MCP tools. */
  static readonly clientToolNotes = [
    "OpenCode V2 Code Mode routing: use a targeted search for the exact excel tool needed, then reuse its discovered schema.",
    "Do not repeatedly dump the full tool catalog. If inventory is requested, list paths and short descriptions once, then look up full schemas only for the tools you will call.",
    "Inside the built-in execute tool, call await tools.excel.<remote_tool_name>(arguments) and return the remote result.",
    "Discover the remote input schema and pass exactly those arguments; do not assume a parameter name from another tool.",
    "If the client's remote tool is named execute, call tools.excel.execute(arguments), not the local execute sandbox with those arguments.",
    "Office code belongs in the code argument of that remote call. Never evaluate Office.js directly inside OpenCode's JavaScript sandbox.",
    "Older OpenCode versions expose prefixed MCP tools directly; use those instead of Code Mode when present.",
  ].join("\n");
  readonly clientToolNotes = OpencodeAdapter.clientToolNotes;

  static readonly capabilities: ProviderCapabilities = {
    streaming: true,
    cancel: true,
    sessionResume: true,
    mcp: true,
    acp: true,
    nativeWebSearch: false,
    explicitToolNames: false,
    credits: false,
    loginKind: "account",
    logout: true,
    liveModel: true,
    liveEffort: true,
  };
  readonly capabilities = OpencodeAdapter.capabilities;

  private readonly sessions = new Map<string, OpencodeSession>();

  private get command(): string {
    return this.env.OPENCODE_BIN ?? "opencode";
  }

  protected probe(): CliProbe {
    return {
      command: this.command,
      versionArgs: ["--version"],
      installHint: "npm install -g opencode-ai",
      loginHint: "opencode auth login",
    };
  }

  async listModels(): Promise<ProviderModelCatalog> {
    const peer = new JsonRpcPeer({
      command: this.command,
      args: ["acp"],
      cwd: ensureWorkspaceDir(),
      env: this.env,
      versionField: true,
      onStderr: (line) => this.runtime.logger.debug(`opencode stderr: ${line}`),
    });
    try {
      await peer.request("initialize", {
        protocolVersion: 1,
        clientCapabilities: {},
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
      }, 15_000);
      const session = await peer.request<{ configOptions?: OpencodeConfigOption[] }>(
        "session/new",
        { cwd: ensureWorkspaceDir(), mcpServers: [] },
        30_000,
      );
      return parseOpencodeModels(session.configOptions);
    } finally {
      await peer.close();
    }
  }

  private auth: OpenCodeAuthApi | null = null;

  /** OpenCode's own sign-in service, started when first needed and shut down when a sign-in ends or nobody uses it. */
  protected openAuthService(): OpenCodeAuthApi {
    if (!this.auth || this.auth.closed) this.auth = new OpenCodeAuthServer({ command: this.command, env: this.env });
    return this.auth;
  }

  /** What OpenCode lets the person change in a session besides model and effort: its mode (build or plan). */
  sessionOptions(providerSessionId: string): ProviderOption[] {
    return acpOptions(this.sessions.get(providerSessionId)?.configOptions as AcpConfigOption[] | undefined, ["mode"]);
  }

  async setSessionOption(providerSessionId: string, optionId: string, value: string): Promise<ProviderOption[]> {
    const session = this.sessions.get(providerSessionId);
    if (!session) throw new Error("OpenCode session is not running");
    const option = this.sessionOptions(providerSessionId).find(item => item.id === optionId);
    if (!option || !option.choices.some(choice => choice.value === value)) throw new Error("OpenCode does not offer that setting.");
    const result = await session.peer.request<{ configOptions?: OpencodeConfigOption[] }>(
      "session/set_config_option", { sessionId: session.acpSessionId, configId: optionId, value }, 10_000);
    if (result?.configOptions) session.configOptions = result.configOptions;
    return this.sessionOptions(providerSessionId);
  }

  /** OpenCode signs out one account at a time. */
  override async logout(accountId?: string): Promise<void> {
    if (!accountId || !/^[A-Za-z0-9][A-Za-z0-9._-]{0,79}$/.test(accountId)) throw new ProviderLoginError("Choose the account to sign out of.");
    // The sign-in service keeps its own copy of the accounts, so it is stopped around the CLI's change and starts fresh for the next list.
    this.auth?.close();
    const result = await execCli(this.command, ["providers", "logout", accountId], this.env, 15_000);
    this.auth?.close();
    if (result.code !== 0) throw new ProviderLoginError("Could not sign out. Try again.");
    this.invalidateDetection();
  }

  async loginAccounts(): Promise<LoginAccount[]> {
    return this.openAuthService().catalog();
  }

  override startLogin(choice?: LoginChoice): ProviderLoginAttempt {
    if (!choice) throw new ProviderLoginError("Choose an account to sign in with.");
    return OpenCodeLoginAttempt.start(this.openAuthService(), choice);
  }

  protected override async checkLogin(command: string): Promise<{ status: ProviderConnectionStatus; hint: string | null }> {
    const result = await execCli(command, ["auth", "list"], this.env, 8_000);
    if (result.code === 0) {
      const count = Number(/(\d+)\s+credentials?\b/i.exec(stripAnsi(`${result.stdout}\n${result.stderr}`))?.[1] ?? 0);
      if (count > 0) return { status: "ready", hint: null };
    }
    if (this.env.ANTHROPIC_API_KEY || this.env.OPENCODE_API_KEY) return { status: "ready", hint: null };
    // OpenCode's free models work without credentials, so the CLI is usable
    // right away; a paid model reports the missing login on the first turn.
    return {
      status: "ready",
      hint: `No OpenCode credentials found. Free models work as they are; run \`opencode auth login\` on ${cliEnvironment()} to use your own subscription.`,
    };
  }

  override async startSession(ctx: StartSessionContext): Promise<string> {
    if (ctx.resumeSessionId && !validNativeId(ctx.resumeSessionId)) throw new Error("Invalid OpenCode resume session ID");
    const cwd = ensureWorkspaceDir();
    const providerSessionId = `opencode-${randomUUID()}`;
    const log = this.runtime.logger;
    // v2 ACP owns a private server and no longer accepts --hostname/--port.
    const clientTools = ctx.mcp.toolNames !== undefined;
    const version = ctx.requestUserInput || clientTools ? (await this.detect()).version ?? "" : "";
    const v2 = isOpencodeV2(version);
    const questionApi = ctx.requestUserInput && !v2 ? { port: await reserveOpencodePort(), password: randomUUID(), cwd, ask: ctx.requestUserInput } : undefined;
    const { dir: configDir, path: configPath } = this.writePermissionConfig(Boolean(questionApi), clientTools && v2);

    let session: OpencodeSession | null = null;
    const peer = new JsonRpcPeer({
      command: this.command,
      args: ["acp", ...(questionApi ? ["--hostname", "127.0.0.1", "--port", String(questionApi.port)] : [])],
      cwd,
      // The permission config is merged on top of the user's own OpenCode
      // config, so their providers, agents and skills keep working.
      env: { ...this.env, OPENCODE_CONFIG: configPath, ...(questionApi ? {
        OPENCODE_ENABLE_QUESTION_TOOL: "1", OPENCODE_SERVER_USERNAME: "excel-partner", OPENCODE_SERVER_PASSWORD: questionApi.password,
      } : {}) },
      versionField: true,
      onStderr: (line) => log.debug(`opencode stderr: ${line}`),
      onNotification: (method, params) => session && this.onNotification(session, method, params),
      onRequest: (method, params) => {
        if (session?.active) session.active.stopWaiting?.();
        return this.onServerRequest(method, params, ctx.authorizeTool);
      },
    });

    try {
      await peer.request("initialize", {
        protocolVersion: 1,
        clientCapabilities: {},
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
      });
      const loaded = await peer.request<{ sessionId?: string; configOptions?: OpencodeConfigOption[] }>(
        ctx.resumeSessionId ? "session/load" : "session/new",
        {
          ...(ctx.resumeSessionId ? { sessionId: ctx.resumeSessionId } : {}),
          cwd,
          mcpServers: [
            {
              type: "stdio",
              name: MCP_SERVER_NAME,
              command: ctx.mcp.command,
              args: ctx.mcp.args,
              env: Object.entries(ctx.mcp.env).map(([name, value]) => ({ name, value })),
            },
          ],
        },
        30_000,
      );
      const created = { ...loaded, sessionId: ctx.resumeSessionId ?? loaded.sessionId };
      if (!created.sessionId) throw new Error("OpenCode returned no session ID");
      session = {
        questionApi,
        peer,
        acpSessionId: created.sessionId,
        configOptions: created.configOptions ?? [],
        model: parseOpencodeModels(created.configOptions).defaultModelId,
        primed: Boolean(ctx.resumeSessionId),
        active: null,
        configDir,
        v2: clientTools && v2,
        reportsUsage: clientTools,
        diagnostic: clientTools ? ctx.onDiagnostic : undefined,
      };
      if (ctx.resumeSessionId && clientTools) {
        session.usage = new CumulativeUsage();
        try { session.usage.next((await this.originalUsage(created.sessionId, clientTools && v2)).usage); }
        catch { session.usageBaselinePending = true; }
      }
      this.sessions.set(providerSessionId, session);
      peer.onExit(() => {
        ctx.onRuntimeClosed?.();
        const current = this.sessions.get(providerSessionId);
        current?.active?.queue.fail(
          new Error(`opencode exited unexpectedly: ${peer.stderrTail.slice(-3).join(" | ")}`),
        );
        this.sessions.delete(providerSessionId);
        fs.rmSync(configDir, { recursive: true, force: true });
      });
      if (ctx.model) await this.selectModel(session, ctx.model);
      return providerSessionId;
    } catch (error) {
      this.sessions.delete(providerSessionId);
      await peer.close();
      fs.rmSync(configDir, { recursive: true, force: true });
      throw error;
    }
  }

  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const { requestId } = input;
    // A prompt that was cancelled is answered by the CLI in a moment; the next one goes after it.
    if (session?.prompting) await Promise.race([session.prompting, new Promise(resolve => setTimeout(resolve, CANCEL_GRACE_MS).unref())]);
    if (!session || !session.peer.alive) {
      yield { type: "error", requestId, code: "no_session", message: "opencode session is not running" };
      yield { type: "turn_finished", requestId, reason: "error" };
      return;
    }

    if (input.command?.name === "skills") {
      yield { type: "session_info", requestId, provider: this.id, providerSessionId: session.acpSessionId, title: session.title };
      yield { type: "turn_started", requestId };
      if (input.signal.aborted) { yield { type: "turn_finished", requestId, reason: "cancelled" }; return; }
      if (input.command.args && input.command.args !== "list") throw new Error("Use /skills or /skills list. Ask OpenCode to use a named skill in your next message.");
      const result = await execCli(this.command, ["api", "GET", "/api/skill", "--standalone"], this.env, 15_000, ensureWorkspaceDir());
      if (result.code !== 0) throw new Error(result.stderr || "Could not list OpenCode skills.");
      const catalog = JSON.parse(result.stdout) as { data: Array<{ name: string; description?: string }> };
      if (!Array.isArray(catalog.data)) throw new Error("OpenCode returned an unsupported skill catalog.");
      yield { type: "message_completed", requestId, text: catalog.data.length ? "Available OpenCode skills:\n\n" + catalog.data.map(skill => `- **${skill.name}** — ${skill.description ?? ""}`).join("\n")
        + "\n\nAsk OpenCode to use a named skill in your next message." : "No OpenCode skills were found." };
      yield { type: "turn_finished", requestId, reason: input.signal.aborted ? "cancelled" : "completed" };
      return;
    }

    const queue = new AsyncQueue<ProviderEvent>();
    const questionController = new AbortController();
    if (session.questionApi) {
      void watchOpencodeStatus({ ...session.questionApi, sessionId: session.acpSessionId,
        signal: AbortSignal.any([input.signal, questionController.signal]),
        onStatus: status => { if (session.active) this.reportProgress(session, status); },
      });
      void watchOpencodeQuestions({ ...session.questionApi, sessionId: session.acpSessionId,
        signal: AbortSignal.any([input.signal, questionController.signal]) }).catch(error => {
        if (input.signal.aborted || questionController.signal.aborted) return;
        session.peer.notify("session/cancel", { sessionId: session.acpSessionId });
        queue.fail(error instanceof Error ? error : new Error(String(error)));
      });
    }
    const active: ActiveTurn = { queue, requestId, text: "", tools: new Map(), finishedTools: new Set() };
    session.active = active;
    yield { type: "session_info", requestId, provider: this.id, providerSessionId: session.acpSessionId,
      title: session.title };
    yield { type: "turn_started", requestId };
    const options = this.sessionOptions(input.providerSessionId);
    if (options.length) yield { type: "provider_options", requestId, options };
    this.reportProgress(session, { state: "waiting" });
    let finished = false;
    let responseTimer: ReturnType<typeof setTimeout> | undefined;
    active.stopWaiting = () => { clearTimeout(responseTimer); responseTimer = undefined; };

    const cancel = () => session.peer.notify("session/cancel", { sessionId: session.acpSessionId });
    const abort = () => {
      if (finished) return;
      finished = true;
      active.stopWaiting?.();
      cancel();
      queue.push({ type: "turn_finished", requestId, reason: "cancelled" });
      queue.end();
      // The CLI confirms a cancel within moments and then takes the next prompt, so it stays open; it is closed only if it never does.
      const stuck = setTimeout(() => void this.closeSession(input.providerSessionId), CANCEL_GRACE_MS);
      stuck.unref();
      void (session.prompting ?? Promise.resolve()).then(() => clearTimeout(stuck));
    };
    input.signal.addEventListener("abort", abort, { once: true });

    // ACP has no system-prompt field, so the instructions travel with the first prompt.
    const userText = buildTurnPrompt(input.prompt, input.selection);
    const systemPrompt = input.systemPrompt ?? SYSTEM_PROMPT;
    const text = input.command ? `/${input.command.name}${input.command.args ? ` ${input.command.args}` : ""}`
      : session.primed || !systemPrompt.trim() ? userText : `${systemPrompt}\n\n---\n\n${userText}`;
    const run = async () => {
      if (input.model && input.model !== session.model) await this.selectModel(session, input.model);
      // OpenCode only offers an effort option for models that have variants.
      const effort = effortConfig(session.configOptions as EffortConfigOption[], input.effort ?? "medium");
      if (effort) {
        const result = await session.peer.request<{ configOptions?: OpencodeConfigOption[] }>(
          "session/set_config_option",
          { sessionId: session.acpSessionId, ...effort },
          10_000,
        );
        if (result?.configOptions) session.configOptions = result.configOptions;
      }
      if (input.signal.aborted) throw new Error("OpenCode request cancelled");
      // A CLI can retry silently for thirty minutes. Bound the initial wait;
      // streaming output and interactive permissions are allowed to continue.
      if (!input.command) session.primed = true;
      // Silent retries are what the short limit is for. A model that is simply thinking for a long time is given longer.
      const waitForResponse = (delay: number, waited: number) => {
        responseTimer = setTimeout(() => {
          if (finished) return;
          waited += delay;
          if (!active.retrying && waited < FIRST_RESPONSE_LIMIT_MS) return waitForResponse(FIRST_RESPONSE_LIMIT_MS - waited, waited);
          finished = true;
          queue.push({ type: "error", requestId, code: "provider_timeout",
            message: `OpenCode received no model response within ${Math.round(waited / 1000)} seconds (${session.model ?? "default model"}). The service may be overloaded. Try again or select another model.` });
          queue.push({ type: "turn_finished", requestId, reason: "error" });
          queue.end();
          cancel();
          void this.closeSession(input.providerSessionId);
        }, delay);
      };
      waitForResponse(FIRST_RESPONSE_RETRY_LIMIT_MS, 0);
      return session.peer.request<{
        stopReason?: string;
        usage?: { inputTokens: number; outputTokens: number; totalTokens: number; cachedReadTokens?: number };
      }>(
        "session/prompt",
        { sessionId: session.acpSessionId, prompt: [{ type: "text", text }, ...(input.images ?? []).map(image => ({ type: "image", mimeType: image.mediaType, data: image.data }))] },
        TURN_IDLE_LIMIT_MS,
        { idle: true },
      );
    };

    const running = run();
    const settled = running.then(() => undefined, () => undefined);
    session.prompting = settled;
    void settled.then(() => { if (session.prompting === settled) session.prompting = undefined; });
    running
      .then(async (response) => {
        if (finished) return;
        finished = true;
        active.stopWaiting?.();
        this.flushText(active);
        await this.reportUsage(session, active);
        const stop = response?.stopReason ?? "end_turn";
        if (stop === "cancelled") {
          queue.push({ type: "turn_finished", requestId, reason: "cancelled" });
        } else if (stop === "end_turn" || stop === "max_turn_requests") {
          if (!active.receivedOutput && active.tools.size === 0) {
            queue.push({ type: "error", requestId, code: "empty_response",
              message: `OpenCode ended without a response (${session.model ?? "default model"}). The model service may be unavailable or overloaded. Try again or select another model.` });
            queue.push({ type: "turn_finished", requestId, reason: "error" });
          } else queue.push({ type: "turn_finished", requestId, reason: "completed" });
        } else {
          queue.push({
            type: "error",
            requestId,
            code: "provider_error",
            message: `opencode stopped the turn (${stop})`,
          });
          queue.push({ type: "turn_finished", requestId, reason: "error" });
        }
        queue.end();
      })
      .catch(async (error) => {
        if (finished) return;
        finished = true;
        active.stopWaiting?.();
        this.flushText(active);
        await this.reportUsage(session, active);
        if (input.signal.aborted) {
          queue.push({ type: "turn_finished", requestId, reason: "cancelled" });
        } else {
          queue.push({
            type: "error",
            requestId,
            code: isAuthError(error) ? "login_required" : "provider_error",
            message: error instanceof Error ? error.message : String(error),
          });
          queue.push({ type: "turn_finished", requestId, reason: "error" });
        }
        queue.end();
      });

    try {
      for await (const event of queue) {
        yield event;
        if (event.type === "turn_finished") break;
      }
    } finally {
      active.stopWaiting?.();
      questionController.abort();
      input.signal.removeEventListener("abort", abort);
      if (session.active === active) session.active = null;
    }
  }

  private async reportUsage(session: OpencodeSession, active: ActiveTurn): Promise<void> {
    if (!session.reportsUsage) return;
    try {
      const snapshot = await this.originalUsage(session.acpSessionId, session.v2);
      session.title = snapshot.title ?? session.title;
      active.queue.push({ type: "session_info", requestId: active.requestId, provider: this.id,
        providerSessionId: session.acpSessionId, title: session.title });
      session.usage ??= new CumulativeUsage();
      const usage = session.usage.next(snapshot.usage);
      if (session.usageBaselinePending) { session.usageBaselinePending = false; return; }
      active.queue.push({ type: "usage_update", requestId: active.requestId, usage });
    } catch (error) {
      // ACP only reports the final model call; using it as the whole turn silently undercounts.
      this.runtime.logger.warn("OpenCode original usage unavailable", { error: String(error) });
    }
  }

  override async cancelTurn(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    if (session?.active) session.peer.notify("session/cancel", { sessionId: session.acpSessionId });
  }

  override async closeSession(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    this.sessions.delete(providerSessionId);
    session?.active?.queue.end();
    await session?.peer.close();
    if (session) fs.rmSync(session.configDir, { recursive: true, force: true });
  }

  override async dispose(): Promise<void> {
    this.auth?.close();
    await Promise.all([...this.sessions.keys()].map((id) => this.closeSession(id)));
  }

  /**
   * Permission config handed to the CLI: every tool has to be approved, and
   * only the Excel MCP tools are pre-approved. Everything else reaches this
   * adapter as a permission request and is approved through the pane.
   */
  private writePermissionConfig(questions = false, v2 = false): { dir: string; path: string } {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-ai-opencode-"));
    const configPath = path.join(dir, "opencode.json");
    fs.writeFileSync(
      configPath,
      JSON.stringify({
        $schema: "https://opencode.ai/config.json",
        ...(v2 ? {
          permissions: [
            { action: "*", resource: "*", effect: "ask" },
            { action: "websearch", resource: "*", effect: "deny" },
            // V2 reaches MCP tools through Code Mode. Nested operations
            // still enforce their own permissions and pane authorization.
            { action: "execute", resource: "*", effect: "allow" },
            { action: `${MCP_SERVER_NAME}_*`, resource: "*", effect: "allow" },
          ],
        } : { permission: {
          "*": "ask",
          websearch: "deny",
          ...(questions ? { question: "allow" } : {}),
          [`${MCP_SERVER_NAME}_*`]: "allow",
        } }),
      }),
      { mode: 0o600 },
    );
    return { dir, path: configPath };
  }

  private async selectModel(session: OpencodeSession, model: string): Promise<void> {
    const result = await session.peer.request<{ configOptions?: OpencodeConfigOption[] }>(
      "session/set_config_option",
      { sessionId: session.acpSessionId, configId: "model", value: model },
      10_000,
    );
    if (result?.configOptions) session.configOptions = result.configOptions;
    session.model = model;
  }

  private reportProgress(session: OpencodeSession, status: OpencodeStatus): void {
    const active = session.active;
    if (!active) return;
    active.retrying = status.state === "retrying";
    active.queue.push({ type: "provider_progress", requestId: active.requestId, state: status.state,
      message: status.state === "retrying"
        ? `OpenCode: ${status.overloaded ? "model service overloaded; retrying" : "model request failed; retrying"}${status.attempt === undefined ? "" : ` (attempt ${status.attempt})`}…`
        : `Waiting for OpenCode (${(session.model ?? "default model").slice(0, 150)})…`,
    });
  }

  private flushText(active: ActiveTurn): void {
    if (!active.text) return;
    active.queue.push({ type: "message_completed", requestId: active.requestId, text: active.text });
    active.text = "";
  }

  private onNotification(session: OpencodeSession, method: string, raw: unknown): void {
    if (method !== "session/update") return;
    const params = (raw ?? {}) as { sessionId?: string; update?: Record<string, any> };
    const update = params.update;
    if (!update) return;
    // ACP exposes retry metadata separately from assistant text/tool events.
    // Forward safe retry state to the pane and optional API diagnostics.
    if (session.active && params.sessionId === session.acpSessionId
      && update.sessionUpdate === "session_info_update" && update._meta && Object.hasOwn(update._meta, "opencode/retry")) {
      const retry = update._meta["opencode/retry"];
      if (retry === null) { this.reportProgress(session, { state: "waiting" }); session.diagnostic?.({ phase: "model_retry_started" }); }
      else if (retry && typeof retry === "object") {
        const error = retry.error;
        this.reportProgress(session, { state: "retrying", attempt: Number.isSafeInteger(retry.attempt) && retry.attempt >= 0 ? retry.attempt : undefined, overloaded: error?.status === 429 || error?.status === 503 });
        session.diagnostic?.({ phase: "model_retry_scheduled",
          ...(Number.isInteger(retry.attempt) && retry.attempt >= 0 ? { attempt: retry.attempt } : {}),
          ...(Number.isInteger(error?.status) && error.status >= 100 && error.status <= 599 ? { status: error.status } : {}),
          ...(typeof error?.type === "string" && /^[a-zA-Z0-9_.-]{1,100}$/.test(error.type) ? { errorType: error.type } : {}),
          ...(typeof retry.nextRetryAt === "string" && Number.isFinite(Date.parse(retry.nextRetryAt))
            ? { nextRetryAt: new Date(retry.nextRetryAt).toISOString() } : {}),
        });
      }
      return;
    }
    if (update.sessionUpdate === "available_commands_update" && params.sessionId === session.acpSessionId) {
      session.commands = normalizeCommands(update.availableCommands);
      return;
    }
    if (update.sessionUpdate === "config_option_update") {
      if (Array.isArray(update.configOptions)) session.configOptions = update.configOptions;
      return;
    }
    const active = session.active;
    if (!active || params.sessionId !== session.acpSessionId) return;
    const { queue, requestId } = active;

    switch (update.sessionUpdate) {
      case "agent_thought_chunk": {
        const text = update.content?.type === "text" ? String(update.content.text ?? "") : "";
        if (text) { active.stopWaiting?.(); queue.push({ type: "thinking_delta", requestId, text }); }
        break;
      }
      case "agent_message_chunk": {
        const text = update.content?.type === "text" ? String(update.content.text ?? "") : "";
        if (text) {
          active.stopWaiting?.();
          active.receivedOutput = true;
          active.text += text;
          queue.push({ type: "message_delta", requestId, text });
        }
        break;
      }
      case "tool_call": {
        active.stopWaiting?.();
        this.flushText(active);
        const callId = String(update.toolCallId);
        const tool = toolNameOf(update);
        active.tools.set(callId, tool);
        queue.push({ type: "tool_started", requestId, callId, tool, args: update.rawInput ?? {} });
        this.maybeFinishTool(active, callId, update);
        break;
      }
      case "tool_call_update": {
        const callId = String(update.toolCallId);
        if (!active.tools.has(callId)) active.tools.set(callId, toolNameOf(update));
        this.maybeFinishTool(active, callId, update);
        break;
      }
      default:
        break;
    }
  }

  private maybeFinishTool(active: ActiveTurn, callId: string, update: Record<string, any>): void {
    const status = update.status;
    if ((status !== "completed" && status !== "failed") || active.finishedTools.has(callId)) return;
    active.finishedTools.add(callId);
    const failed = status === "failed";
    active.queue.push({
      type: "tool_finished",
      requestId: active.requestId,
      callId,
      tool: active.tools.get(callId) ?? "unknown",
      ok: !failed,
      error: failed ? contentText(update).slice(0, 500) || "tool call failed" : null,
    });
  }

  private async onServerRequest(method: string, params: unknown, authorize?: StartSessionContext["authorizeTool"]): Promise<unknown> {
    if (method !== "session/request_permission") throw new Error(`unsupported client method: ${method}`);
    const request = (params ?? {}) as { toolCall?: Record<string, any>; options?: Array<{ optionId: string; kind?: string }> };
    if (toolNameOf(request.toolCall ?? {}, true).isExcel) return acpPermissionOutcome(params);
    // Fail closed when disconnected or when the pane cannot answer.
    let approved = false;
    try { approved = await authorize?.(request.toolCall ?? {}) ?? false; } catch { /* reject below */ }
    const kind = approved ? "allow_once" : "reject_once";
    const choice = request.options?.find(option => option.kind === kind);
    return choice ? { outcome: { outcome: "selected", optionId: choice.optionId } } : { outcome: { outcome: "cancelled" } };
  }
}

function isAuthError(error: unknown): boolean {
  return /auth|login|401|403|token|credential/i.test(error instanceof Error ? error.message : String(error));
}

function stripAnsi(value: string): string {
  return value.replace(/\u001b\[[0-9;]*m/g, "");
}
