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

interface CodexSession {
  peer: JsonRpcPeer;
  threadId: string;
  tokenTotals?: import("../../protocol/src/index.js").TokenUsage;
  /** Sink for the turn currently running, if any. */
  active: { queue: AsyncQueue<ProviderEvent>; requestId: string; turnId: string | null; interrupting: boolean; tokenBaseline?: import("../../protocol/src/index.js").TokenUsage } | null;
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
    const command = process.env.CODEX_BIN ?? "codex";
    const cwd = ensureWorkspaceDir();
    const providerSessionId = `codex-${randomUUID()}`;
    const log = this.runtime.logger;

    let session: CodexSession | null = null;
    const peer = new JsonRpcPeer({
      command,
      args: ["app-server"],
      cwd,
      env: this.env,
      onStderr: (line) => log.debug(`codex stderr: ${line}`),
      onNotification: (method, params) => session && this.onNotification(session, method, params),
      onRequest: (method, params) => this.onServerRequest(method, params),
    });

    try {
      await peer.request("initialize", {
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
        capabilities: null,
      });
      peer.notify("initialized");

      const started = await peer.request<{ thread: { id: string } }>("thread/start", {
        cwd,
        ...(ctx.model ? { model: ctx.model } : {}),
        approvalPolicy: "never",
        sandbox: "read-only",
        ephemeral: true,
        developerInstructions: SYSTEM_PROMPT,
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
          web_search: "disabled",
        },
      });

      session = { peer, threadId: started.thread.id, active: null };
      this.sessions.set(providerSessionId, session);
      peer.onExit(() => {
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
    yield { type: "turn_started", requestId };

    const onAbort = () => void this.interrupt(session);
    input.signal.addEventListener("abort", onAbort, { once: true });

    session.peer
      .request<{ turn: TurnLike }>("turn/start", {
        threadId: session.threadId,
        input: [{ type: "text", text: buildTurnPrompt(input.prompt, input.selection), text_elements: [] }],
        effort: input.effort ?? "medium",
      })
      .then((response) => {
        if (session.active?.queue === queue) {
          session.active.turnId = response.turn.id;
          if (input.signal.aborted) void this.interrupt(session);
        }
      })
      .catch((error) => {
        queue.push({
          type: "error",
          requestId,
          code: "provider_error",
          message: error instanceof Error ? error.message : String(error),
        });
        queue.push({ type: "turn_finished", requestId, reason: "error" });
        queue.end();
      });

    try {
      for await (const event of queue) {
        yield event;
        if (event.type === "turn_finished") break;
      }
    } finally {
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
    const active = session.active;
    if (!active || params.threadId !== session.threadId) return;
    const { queue, requestId } = active;

    switch (method) {
      case "thread/tokenUsage/updated": {
        if (params.turnId && active.turnId && params.turnId !== active.turnId) break;
        const total = params.tokenUsage?.total;
        if (!total || ![total.inputTokens, total.outputTokens, total.totalTokens].every((n) => Number.isSafeInteger(n) && n >= 0)) break;
        active.tokenBaseline ??= session.tokenTotals ?? { inputTokens: 0, outputTokens: 0, cachedInputTokens: 0, totalTokens: 0 };
        const current = { inputTokens: total.inputTokens, outputTokens: total.outputTokens,
          cachedInputTokens: total.cachedInputTokens ?? 0, totalTokens: total.totalTokens };
        const baseline = active.tokenBaseline;
        queue.push({ type: "usage_update", requestId, usage: {
          inputTokens: Math.max(0, current.inputTokens - baseline.inputTokens),
          outputTokens: Math.max(0, current.outputTokens - baseline.outputTokens),
          cachedInputTokens: Math.max(0, current.cachedInputTokens - baseline.cachedInputTokens),
          totalTokens: Math.max(0, current.totalTokens - baseline.totalTokens),
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

  private onServerRequest(method: string, params: unknown): unknown {
    switch (method) {
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
