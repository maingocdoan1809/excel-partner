import { randomUUID } from "node:crypto";
import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderEvent,
  ProviderId,
} from "../../protocol/src/index.js";
import { AsyncQueue } from "./async-queue.js";
import { BaseCliAdapter, cliEnvironment, execCli, type CliProbe } from "./base.js";
import { JsonRpcPeer } from "./jsonrpc.js";
import { buildTurnPrompt, SYSTEM_PROMPT } from "./prompt.js";
import { MCP_SERVER_NAME, type ProviderModelCatalog, type SendTurnInput, type StartSessionContext } from "./types.js";
import { ensureWorkspaceDir } from "./workspace.js";
import { codexTokenBaseline, nativeTitle, validNativeId } from "./session-metadata.js";
import { answerCodexQuestions, type RequestUserInput } from "./user-input.js";
import { CODEX_COMMANDS } from "./codex-commands.js";

interface CodexSession {
  model?: string;
  mode?: "plan" | "default";
  questions: Map<number | string, AbortController>;
  peer: JsonRpcPeer;
  threadId: string;
  title?: string;
  baselinePending?: boolean;
  tokenTotals?: import("../../protocol/src/index.js").TokenUsage;
  /** Sink for the turn currently running, if any. */
  active: { queue: AsyncQueue<ProviderEvent>; requestId: string; turnId: string | null; interrupting: boolean; tokenBaseline?: import("../../protocol/src/index.js").TokenUsage; usageUnavailable?: boolean } | null;
}

interface TurnLike {
  id: string;
  status: "completed" | "interrupted" | "failed" | "inProgress";
  error?: { message?: string } | null;
}

/**
 * Codex CLI via `codex app-server` (stdio, newline-delimited JSON-RPC).
 *
 * Verified against codex-cli 0.162.1:
 *   initialize -> thread/start (config.mcp_servers.excel) -> turn/start
 *   notifications: item/agentMessage/delta, item/started|completed (mcpToolCall),
 *                  turn/completed; cancel = turn/interrupt.
 */
export class CodexAdapter extends BaseCliAdapter {
  async listCommands() { return CODEX_COMMANDS; }
  async sessionTitle(id: string): Promise<string | undefined> {
    const session = [...this.sessions.values()].find(s => s.threadId === id);
    if (!session) return undefined;
    try {
      const result = await session.peer.request<{ thread: { name?: string | null } }>("thread/read", { threadId: id, includeTurns: false }, 5000);
      session.title = nativeTitle(result.thread.name) ?? session.title;
    } catch { /* The thread can be busy or the CLI may not implement thread/read. */ }
    return session.title;
  }
  async listModels(): Promise<ProviderModelCatalog> {
    const peer = new JsonRpcPeer({ command: process.env.CODEX_BIN ?? "codex", args: ["app-server"], env: this.env });
    try {
      await peer.request("initialize", {
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
        capabilities: null,
      }, 15_000);
      peer.notify("initialized");
      const result = await peer.request<{
        data?: Array<{ id?: string; model?: string; displayName?: string; hidden?: boolean; isDefault?: boolean }>;
      }>("model/list", { limit: 100 }, 20_000);
      const visible = (result.data ?? []).filter((item) => !item.hidden && (item.model || item.id));
      return {
        models: visible.map((item) => ({ id: (item.model || item.id)!, name: item.displayName || (item.model || item.id)! })),
        defaultModelId: visible.find((item) => item.isDefault)?.model ?? null,
      };
    } finally {
      await peer.close();
    }
  }
  readonly id: ProviderId = "codex";
  readonly capabilities: ProviderCapabilities = {
    streaming: true,
    cancel: true,
    sessionResume: true,
    mcp: true,
  };

  /** Lowest Codex version this adapter has been tested with. */
  static readonly MIN_VERSION = "0.162.1";

  protected override get minVersion(): string {
    return CodexAdapter.MIN_VERSION;
  }

  private readonly sessions = new Map<string, CodexSession>();

  protected probe(): CliProbe {
    return {
      command: process.env.CODEX_BIN ?? "codex",
      versionArgs: ["--version"],
      installHint: "npm install -g @openai/codex",
      loginHint: "codex login",
    };
  }

  protected override async checkLogin(command: string): Promise<{ status: ProviderConnectionStatus; hint: string | null }> {
    const result = await execCli(command, ["login", "status"], this.env);
    const text = `${result.stdout}\n${result.stderr}`;
    if (result.code === 0 && /logged in/i.test(text)) return { status: "ready", hint: null };
    return { status: "login_required", hint: `Run \`codex login\` on ${cliEnvironment()}, then retry.` };
  }

  override async startSession(ctx: StartSessionContext): Promise<string> {
    if (ctx.resumeSessionId && !validNativeId(ctx.resumeSessionId)) throw new Error("Invalid Codex resume thread ID");
    const command = process.env.CODEX_BIN ?? "codex";
    const cwd = ensureWorkspaceDir();
    const providerSessionId = `codex-${randomUUID()}`;
    const log = this.runtime.logger;

    let session: CodexSession | null = null;
    const peer = new JsonRpcPeer({
      command,
      args: ["app-server", "-c", `web_search="${ctx.mcp.env.EXCEL_WEB_SEARCH === "1" ? "live" : "disabled"}"`],
      cwd,
      env: this.env,
      onStderr: (line) => log.debug(`codex stderr: ${line}`),
      onNotification: (method, params) => session && this.onNotification(session, method, params),
      onRequest: (method, params, id) => this.onServerRequest(method, params, ctx.requestUserInput, session, id),
    });

    try {
      await peer.request("initialize", {
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
        capabilities: { experimentalApi: true },
      });
      peer.notify("initialized");

      const started = await peer.request<{ model?: string; thread: { id: string; name?: string | null; path?: string | null } }>(ctx.resumeSessionId ? "thread/resume" : "thread/start", {
        ...(ctx.resumeSessionId ? { threadId: ctx.resumeSessionId } : { ephemeral: false }),
        cwd,
        ...(ctx.model ? { model: ctx.model } : {}),
        approvalPolicy: "never",
        sandbox: "read-only",
        developerInstructions: [
          "Current session tool configuration supersedes Excel-only and no-web restrictions from earlier turns.",
          `Native web search: ${ctx.mcp.env.EXCEL_WEB_SEARCH === "1" ? "enabled; use the provider's built-in web search for web requests" : "disabled"}.`,
          `Playwright browser: ${ctx.mcp.env.EXCEL_PLAYWRIGHT === "1" ? "enabled through the Excel MCP playwright_browser tool" : "disabled"}.`,
          "Being embedded in Excel does not prohibit enabled web tools. Determine availability from the current tools and their results, not earlier assistant replies.",
          ctx.systemPrompt ?? SYSTEM_PROMPT,
        ].filter(Boolean).join("\n"),
        config: {
          mcp_servers: {
            [MCP_SERVER_NAME]: {
              command: ctx.mcp.command,
              args: ctx.mcp.args,
              env: ctx.mcp.env,
              startup_timeout_sec: 20,
              default_tools_approval_mode: "approve",
              tool_timeout_sec: 120,
            },
          },
          web_search: ctx.mcp.env.EXCEL_WEB_SEARCH === "1" ? "live" : "disabled",
        },
      });

      if (ctx.resumeSessionId && started.thread.id !== ctx.resumeSessionId) throw new Error("Codex returned a different thread while resuming");
      const baseline = ctx.resumeSessionId ? await codexTokenBaseline(started.thread.path) : undefined;
      session = { peer, threadId: started.thread.id, title: nativeTitle(started.thread.name), model: ctx.model ?? started.model, active: null, questions: new Map(),
        tokenTotals: baseline, baselinePending: Boolean(ctx.resumeSessionId && !baseline) };
      this.sessions.set(providerSessionId, session);
      peer.onExit(() => {
        ctx.onRuntimeClosed?.();
        for (const question of session?.questions.values() ?? []) question.abort();
        const current = this.sessions.get(providerSessionId);
        if (current?.active) {
          current.active.queue.fail(
            new Error(`codex app-server exited unexpectedly: ${peer.stderrTail.slice(-3).join(" | ")}`),
          );
        }
        this.sessions.delete(providerSessionId);
      });
      return providerSessionId;
    } catch (error) {
      await peer.close();
      throw error;
    }
  }

  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const { requestId } = input;
    if (!session || !session.peer.alive) {
      yield { type: "error", requestId, code: "no_session", message: "codex session is not running" };
      yield { type: "turn_finished", requestId, reason: "error" };
      return;
    }

    const queue = new AsyncQueue<ProviderEvent>();
    session.active = { queue, requestId, turnId: null, interrupting: false };
    yield { type: "session_info", requestId, provider: this.id, providerSessionId: session.threadId, title: session.title };
    yield { type: "turn_started", requestId };

    const onAbort = () => void this.interrupt(session);
    input.signal.addEventListener("abort", onAbort, { once: true });

    const run = async (): Promise<{ turn: TurnLike } | null> => {
      if (input.command) {
        const { name, args } = input.command;
        const text = (value: string) => {
          queue.push({ type: "message_completed", requestId, text: value });
          queue.push({ type: "turn_finished", requestId, reason: input.signal.aborted ? "cancelled" : "completed" });
          queue.end();
          return null;
        };
        input.signal.throwIfAborted();
        if (name === "compact") {
          await session.peer.request("thread/compact/start", { threadId: session.threadId });
          return null; // Completion and usage arrive through the normal turn notifications.
        }
        if (name === "plan") {
          session.mode = args === "off" ? "default" : args === "on" || args ? "plan" : session.mode === "plan" ? "default" : "plan";
          if (!args || args === "on" || args === "off") return text(`Codex ${session.mode === "plan" ? "Plan" : "Default"} mode enabled for the next message.`);
          if (!session.model) throw new Error("Choose a model before starting a planning task.");
          return session.peer.request("turn/start", { threadId: session.threadId,
            input: [{ type: "text", text: args, text_elements: [] }],
            collaborationMode: { mode: session.mode, settings: { model: session.model, reasoning_effort: input.effort ?? "medium", developer_instructions: null } } });
        }
        if (name === "goal") {
          const method = !args ? "thread/goal/get" : args === "clear" ? "thread/goal/clear" : "thread/goal/set";
          const result = await session.peer.request(method, { threadId: session.threadId, ...(args && args !== "clear" ? { objective: args, status: "active" } : {}) });
          return text(`Goal:\n\n\`\`\`json\n${JSON.stringify(result, null, 2)}\n\`\`\``);
        }
        if (name === "fork") {
          const result = await session.peer.request<{ thread: { id: string; path?: string | null; name?: string | null } }>("thread/fork", { threadId: session.threadId, excludeTurns: true });
          session.threadId = result.thread.id;
          session.title = nativeTitle(result.thread.name);
          session.tokenTotals = await codexTokenBaseline(result.thread.path);
          session.baselinePending = !session.tokenTotals;
          if (args) { await session.peer.request("thread/name/set", { threadId: session.threadId, name: args }); session.title = args; }
          queue.push({ type: "session_info", requestId, provider: this.id, providerSessionId: session.threadId, title: session.title });
          return text("Forked this conversation. Your next message continues in the new session.");
        }
        if (name === "review") {
          const target = !args || args === "uncommitted" ? { type: "uncommittedChanges" }
            : args.startsWith("branch ") ? { type: "baseBranch", branch: args.slice(7).trim() }
            : args.startsWith("commit ") ? { type: "commit", sha: args.slice(7).trim() }
            : { type: "custom", instructions: args };
          return session.peer.request("review/start", { threadId: session.threadId, target, delivery: "inline" });
        }
        if (name === "rename") {
          if (!args) throw new Error("Usage: /rename <title>");
          await session.peer.request("thread/name/set", { threadId: session.threadId, name: args });
          session.title = args;
          queue.push({ type: "session_info", requestId, provider: this.id, providerSessionId: session.threadId, title: args });
          return text(`Renamed this session to ${args}.`);
        }
        if (name === "skills") {
          const result = await session.peer.request<{ data: Array<{ skills: Array<{ name: string; description: string; path: string; enabled: boolean }>; errors?: unknown[] }> }>(
            "skills/list", { cwds: [ensureWorkspaceDir()], forceReload: true });
          const skills = result.data.flatMap(entry => entry.skills).filter(skill => skill.enabled);
          if (!args || args === "list") return text(skills.length ? "Available Codex skills:\n\n" + skills.map(skill => `- **${skill.name}** — ${skill.description}`).join("\n")
            + "\n\nUse `/skills <name> <request>` to apply a skill." : "No enabled Codex skills were found.");
          const [, skillName, task] = /^(\S+)(?:\s+([\s\S]*))?$/.exec(args)!;
          const skill = skills.find(item => item.name === skillName);
          if (!skill) throw new Error(`Unknown Codex skill: ${skillName}. Use /skills to see available skills.`);
          if (!task) return text(`**${skill.name}** — ${skill.description}\n\nUse \`/skills ${skill.name} <request>\` to apply this skill.`);
          return session.peer.request("turn/start", { threadId: session.threadId, effort: input.effort ?? "medium",
            input: [{ type: "text", text: `$${skill.name} ${task}`, text_elements: [] }, { type: "skill", name: skill.name, path: skill.path }] });
        }
        const methods: Record<string, [string, unknown]> = {
          mcp: ["mcpServerStatus/list", {}], apps: ["app/list", { threadId: session.threadId }],
          ps: ["thread/backgroundTerminals/list", { threadId: session.threadId }],
          stop: ["thread/backgroundTerminals/clean", { threadId: session.threadId }],
          status: ["thread/read", { threadId: session.threadId, includeTurns: false }],
          usage: ["account/usage/read", {}],
          hooks: ["hooks/list", { cwds: [ensureWorkspaceDir()] }],
        };
        if (name === "pwd") return text(ensureWorkspaceDir());
        const method = methods[name];
        if (!method) throw new Error(`/${name} is not available in the add-in.`);
        const result = await session.peer.request(method[0], method[1]);
        return text(`/${name}\n\n\`\`\`json\n${JSON.stringify(result, null, 2)}\n\`\`\``);
      }
      return session.peer.request<{ turn: TurnLike }>("turn/start", {
        threadId: session.threadId,
        input: [{ type: "text", text: buildTurnPrompt(input.prompt, input.selection), text_elements: [] },
          ...(input.images ?? []).map(image => ({ type: "image", image: { type: "inline", url: `data:${image.mediaType};base64,${image.data}` }, detail: null }))],
        effort: input.effort ?? "medium",
        ...(session.mode && session.model ? { collaborationMode: { mode: session.mode,
          settings: { model: session.model, reasoning_effort: input.effort ?? "medium", developer_instructions: null } } } : {}),
      });
    };
    run()
      .then((response) => {
        if (response?.turn && session.active?.queue === queue) {
          session.active.turnId = response.turn.id;
          if (input.signal.aborted) void this.interrupt(session);
        }
      })
      .catch((error) => {
        if (!input.signal.aborted) queue.push({
          type: "error",
          requestId,
          code: "provider_error",
          message: error instanceof Error ? error.message : String(error),
        });
        queue.push({ type: "turn_finished", requestId, reason: input.signal.aborted ? "cancelled" : "error" });
        queue.end();
      });

    try {
      for await (const event of queue) {
        if (event.type === "turn_finished") yield { type: "session_info", requestId, provider: this.id,
          providerSessionId: session.threadId, title: await this.sessionTitle(session.threadId) };
        yield event;
        if (event.type === "turn_finished") break;
      }
    } finally {
      for (const question of session.questions.values()) question.abort();
      input.signal.removeEventListener("abort", onAbort);
      if (session.active?.queue === queue) session.active = null;
    }
  }

  override async cancelTurn(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    if (session) await this.interrupt(session);
  }

  override async closeSession(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    this.sessions.delete(providerSessionId);
    for (const question of session?.questions.values() ?? []) question.abort();
    session?.active?.queue.end();
    await session?.peer.close();
  }

  override async dispose(): Promise<void> {
    await Promise.all([...this.sessions.keys()].map((id) => this.closeSession(id)));
  }

  private async interrupt(session: CodexSession): Promise<void> {
    const active = session.active;
    if (!active || active.interrupting) return;
    active.interrupting = true;
    for (const question of session.questions.values()) question.abort();
    // The turn id arrives with the turn/start response or turn/started
    // notification; wait briefly if the user cancels immediately.
    for (let i = 0; i < 40 && !active.turnId && session.active === active; i++) {
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    if (!active.turnId) return;
    try {
      await session.peer.request("turn/interrupt", { threadId: session.threadId, turnId: active.turnId }, 10_000);
    } catch (error) {
      this.runtime.logger.warn("codex turn/interrupt failed", { error: String(error) });
    }
  }

  private onNotification(session: CodexSession, method: string, raw: unknown): void {
    const params = (raw ?? {}) as Record<string, any>;
    if (params.threadId !== session.threadId) return;
    if (method === "serverRequest/resolved") {
      session.questions.get(params.requestId)?.abort();
      return;
    }
    if (method === "turn/completed") for (const question of session.questions.values()) question.abort();
    if (method === "thread/name/updated") session.title = nativeTitle(params.threadName);
    const active = session.active;
    if (!active) return;
    const { queue, requestId } = active;

    switch (method) {
      case "thread/name/updated":
        queue.push({ type: "session_info", requestId, provider: this.id, providerSessionId: session.threadId, title: session.title });
        break;
      case "thread/tokenUsage/updated": {
        if (params.turnId && active.turnId && params.turnId !== active.turnId) break;
        const total = params.tokenUsage?.total;
        if (!total || ![total.inputTokens, total.outputTokens, total.totalTokens].every((n) => Number.isSafeInteger(n) && n >= 0)) break;
        active.tokenBaseline ??= session.tokenTotals ?? { inputTokens: 0, outputTokens: 0, cachedInputTokens: 0, totalTokens: 0,
          cacheWriteTokens: 0, uncachedInputTokens: 0, reasoningTokens: 0 };
        const current = { inputTokens: total.inputTokens, outputTokens: total.outputTokens,
          cachedInputTokens: total.cachedInputTokens ?? 0, totalTokens: total.totalTokens,
          ...(Number.isSafeInteger(total.cacheWriteInputTokens) && total.cacheWriteInputTokens >= 0 && (total.cachedInputTokens ?? 0) + total.cacheWriteInputTokens <= total.inputTokens
            ? { cacheWriteTokens: total.cacheWriteInputTokens as number, uncachedInputTokens: total.inputTokens - (total.cachedInputTokens ?? 0) - total.cacheWriteInputTokens } : {}),
          ...(Number.isSafeInteger(total.reasoningOutputTokens) && total.reasoningOutputTokens >= 0 ? { reasoningTokens: total.reasoningOutputTokens as number } : {}) };
        if (session.baselinePending) {
          session.tokenTotals = current;
          active.tokenBaseline = current;
          active.usageUnavailable = true;
          session.baselinePending = false;
          break;
        }
        if (active.usageUnavailable) { session.tokenTotals = current; break; }
        const baseline = active.tokenBaseline;
        queue.push({ type: "usage_update", requestId, usage: {
          inputTokens: Math.max(0, current.inputTokens - baseline.inputTokens),
          outputTokens: Math.max(0, current.outputTokens - baseline.outputTokens),
          cachedInputTokens: Math.max(0, current.cachedInputTokens - baseline.cachedInputTokens),
          totalTokens: Math.max(0, current.totalTokens - baseline.totalTokens),
          source: "codex_thread", providerSessionId: session.threadId,
          ...(current.cacheWriteTokens !== undefined && baseline.cacheWriteTokens !== undefined && current.uncachedInputTokens !== undefined && baseline.uncachedInputTokens !== undefined
            ? { cacheWriteTokens: Math.max(0, current.cacheWriteTokens - baseline.cacheWriteTokens),
              uncachedInputTokens: Math.max(0, current.uncachedInputTokens - baseline.uncachedInputTokens) } : {}),
          ...(current.reasoningTokens !== undefined && baseline.reasoningTokens !== undefined
            ? { reasoningTokens: Math.max(0, current.reasoningTokens - baseline.reasoningTokens) } : {}),
        } });
        session.tokenTotals = current;
        break;
      }
      case "turn/started":
        active.turnId = params.turn?.id ?? active.turnId;
        break;
      case "item/agentMessage/delta":
        if (params.delta) queue.push({ type: "message_delta", requestId, text: String(params.delta) });
        break;
      case "item/started": {
        const item = params.item;
        if (item?.type === "mcpToolCall") {
          queue.push({ type: "tool_started", requestId, callId: item.id, tool: item.tool, args: item.arguments });
        } else if (item?.type === "webSearch") {
          queue.push({ type: "tool_started", requestId, callId: item.id, tool: "web_search", args: { query: item.query, action: item.action } });
        }
        break;
      }
      case "item/completed": {
        const item = params.item;
        if (item?.type === "agentMessage") {
          queue.push({ type: "message_completed", requestId, text: String(item.text ?? "") });
        } else if (item?.type === "mcpToolCall") {
          const toolError = item.error?.message ?? (item.result?.isError ? toolErrorText(item.result) : null);
          queue.push({
            type: "tool_finished",
            requestId,
            callId: item.id,
            tool: item.tool,
            ok: item.status === "completed" && !toolError,
            error: item.status === "failed" || toolError ? (toolError ?? "tool call failed") : null,
          });
        } else if (item?.type === "webSearch") {
          queue.push({ type: "tool_finished", requestId, callId: item.id, tool: "web_search", ok: true, error: null });
        } else if (item?.type === "commandExecution" || item?.type === "fileChange") {
          // The sandbox is read-only and approvals are denied; surface the attempt.
          this.runtime.logger.warn("codex attempted a non-Excel action", { type: item.type });
        }
        break;
      }
      case "error":
        if (!params.willRetry) {
          queue.push({
            type: "error",
            requestId,
            code: "provider_error",
            message: String(params.error?.message ?? "codex reported an error"),
          });
        }
        break;
      case "turn/completed": {
        queue.push({ type: "session_info", requestId, provider: this.id, providerSessionId: session.threadId, title: session.title });
        const turn = params.turn as TurnLike;
        if (turn.status === "failed") {
          queue.push({
            type: "error",
            requestId,
            code: "provider_error",
            message: String(turn.error?.message ?? "codex turn failed"),
          });
        }
        queue.push({
          type: "turn_finished",
          requestId,
          reason: turn.status === "interrupted" ? "cancelled" : turn.status === "failed" ? "error" : "completed",
        });
        queue.end();
        break;
      }
    }
  }

  private onServerRequest(method: string, params: unknown, ask?: RequestUserInput, session?: CodexSession | null, id?: number | string): unknown {
    switch (method) {
      case "item/tool/requestUserInput": {
        const request = params as Record<string, any>;
        if (!session?.active || session.active.interrupting || request.threadId !== session.threadId
          || (session.active.turnId && request.turnId !== session.active.turnId)) return { answers: {} };
        const controller = new AbortController();
        if (id !== undefined) session.questions.set(id, controller);
        return answerCodexQuestions(request, ask, controller.signal).finally(() => {
          if (id !== undefined) session.questions.delete(id);
        });
      }
      case "mcpServer/elicitation/request": {
        // Tool-call approval for our own Excel server is auto-accepted; anything else is declined.
        const server = (params as { serverName?: string } | null)?.serverName;
        return server === MCP_SERVER_NAME
          ? { action: "accept", content: null, _meta: null }
          : { action: "decline", content: null, _meta: null };
      }
      case "item/commandExecution/requestApproval":
      case "item/fileChange/requestApproval":
        return { decision: "decline" };
      case "item/permissions/requestApproval":
        return { permissions: {}, scope: "turn" };
      case "execCommandApproval":
      case "applyPatchApproval":
        return { decision: "denied" };
      default:
        throw new Error(`unsupported server request: ${method}`);
    }
  }
}

function toolErrorText(result: { content?: Array<{ text?: string }> }): string {
  return result.content?.map((c) => c.text ?? "").join(" ").slice(0, 500) || "tool returned an error";
}
