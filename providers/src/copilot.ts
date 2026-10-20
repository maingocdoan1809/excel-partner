import { randomUUID } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { exposedToolNames } from "../../protocol/src/index.js";
import { TokenUsageSchema } from "../../protocol/src/usage.js";
import { CumulativeUsage, parseCopilotCredits } from "./usage.js";
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
import { effortConfig, type EffortConfigOption } from "./effort.js";

interface CopilotSession {
  peer: JsonRpcPeer;
  /** Directory holding the 0600 MCP config (contains the session token). */
  configDir: string;
  acpSessionId: string;
  /** True once the system prompt has been delivered with the first prompt. */
  primed: boolean;
  active: ActiveTurn | null;
  usage?: CumulativeUsage;
  supportsUsageCommand?: boolean;
  usageReport?: string;
  creditTotal?: number;
}

interface ActiveTurn {
  queue: AsyncQueue<ProviderEvent>;
  requestId: string;
  /** Assistant text accumulated since the last tool call / message boundary. */
  text: string;
  tools: Map<string, string>;
  finishedTools: Set<string>;
}

/**
 * GitHub Copilot CLI via ACP (`copilot --acp`, JSON-RPC over stdio).
 *
 * Verified against Copilot CLI 1.0.95:
 *   initialize -> session/new (mcpServers: [excel]) -> session/prompt
 *   notifications: session/update (agent_message_chunk, tool_call,
 *   tool_call_update); cancel = session/cancel notification.
 */
export class CopilotAdapter extends BaseCliAdapter {
  async listModels(): Promise<ProviderModelCatalog> {
    const peer = new JsonRpcPeer({
      command: this.command,
      args: ["--acp", "--no-auto-update", "--no-ask-user", "--no-custom-instructions", "--no-remote", "--disable-builtin-mcps"],
      cwd: ensureWorkspaceDir(),
      env: this.env,
      versionField: true,
    });
    try {
      await peer.request("initialize", {
        protocolVersion: 1,
        clientCapabilities: {},
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
      }, 15_000);
      const session = await peer.request<{
        models?: { availableModels?: Array<{ modelId?: string; name?: string; _meta?: { copilotEnablement?: string } }>; currentModelId?: string };
      }>("session/new", { cwd: ensureWorkspaceDir(), mcpServers: [] }, 30_000);
      const seen = new Set<string>();
      const models = (session.models?.availableModels ?? []).flatMap((item) => {
        if (!item.modelId || item._meta?.copilotEnablement === "disabled" || seen.has(item.modelId)) return [];
        seen.add(item.modelId);
        return [{ id: item.modelId, name: item.name || item.modelId }];
      });
      return { models, defaultModelId: session.models?.currentModelId ?? null };
    } finally {
      await peer.close();
    }
  }
  readonly id: ProviderId = "copilot";
  readonly capabilities: ProviderCapabilities = {
    streaming: true,
    cancel: true,
    sessionResume: true,
    mcp: true,
    acp: true,
  };

  /** Lowest Copilot CLI version this adapter has been tested with. */
  static readonly MIN_VERSION = "1.0.95";

  protected override get minVersion(): string {
    return CopilotAdapter.MIN_VERSION;
  }

  private readonly sessions = new Map<string, CopilotSession>();

  private get command(): string {
    return process.env.COPILOT_BIN ?? "copilot";
  }

  protected probe(): CliProbe {
    return {
      command: this.command,
      versionArgs: ["--version"],
      installHint: "npm install -g @github/copilot",
      loginHint: "copilot login",
    };
  }

  protected override async checkLogin(
    _command: string,
  ): Promise<{ status: ProviderConnectionStatus; hint: string | null }> {
    // The Copilot CLI has no `auth status`; its login is either a stored
    // token or a GitHub token in the environment / `gh`. Accept those signals
    // and otherwise let the first turn surface an auth error precisely.
    const env = this.env;
    if (env.COPILOT_GITHUB_TOKEN || env.GH_TOKEN || env.GITHUB_TOKEN) return { status: "ready", hint: null };
    const gh = await execCli("gh", ["auth", "status"], env, 8_000);
    if (gh.code === 0) return { status: "ready", hint: null };
    const fs = await import("node:fs");
    const os = await import("node:os");
    const path = await import("node:path");
    const configDir = env.COPILOT_HOME ?? path.join(os.homedir(), ".copilot");
    if (fs.existsSync(path.join(configDir, "config.json")) || fs.existsSync(path.join(configDir, "apps.json"))) {
      return { status: "ready", hint: null };
    }
    return { status: "login_required", hint: `Run \`copilot login\` on ${cliEnvironment()}, then retry.` };
  }

  override async startSession(ctx: StartSessionContext): Promise<string> {
    const cwd = ensureWorkspaceDir();
    const providerSessionId = `copilot-${randomUUID()}`;
    const log = this.runtime.logger;

    // Copilot's ACP server rejects stdio MCP servers passed in `session/new`
    // ("Rejecting non-http/sse MCP server"), so the Excel server is handed over
    // through `--additional-mcp-config @file` instead. The file is 0600
    // because it carries the per-session bridge token.
    const configDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-ai-copilot-"));
    const mcpConfigPath = path.join(configDir, "mcp.json");
    fs.writeFileSync(
      mcpConfigPath,
      JSON.stringify({
        mcpServers: {
          [MCP_SERVER_NAME]: {
            type: "local",
            command: ctx.mcp.command,
            args: ctx.mcp.args,
            env: ctx.mcp.env,
            tools: ["*"],
          },
        },
      }),
      { mode: 0o600 },
    );

    let session: CopilotSession | null = null;
    let supportsUsageCommand = false;
    const peer = new JsonRpcPeer({
      command: this.command,
      args: [
        "--acp",
        ...(ctx.model ? ["--model", ctx.model] : []),
        "--additional-mcp-config",
        `@${mcpConfigPath}`,
        "--no-auto-update",
        "--no-ask-user",
        "--no-custom-instructions",
        "--no-remote",
        "--disable-builtin-mcps",
        // Only the Excel MCP tools are visible to the model: no shell, file or web tools.
        "--available-tools",
        ...exposedToolNames(ctx.mcp.env.EXCEL_ALLOW_OFFICE_JS === "1").map((name) => `${MCP_SERVER_NAME}-${name}`),
        "--allow-tool",
        MCP_SERVER_NAME,
      ],
      cwd,
      env: this.env,
      versionField: true,
      onStderr: (line) => log.debug(`copilot stderr: ${line}`),
      onNotification: (method, params) => {
        const update = (params as { update?: { sessionUpdate?: string; availableCommands?: Array<{ name: string }> } })?.update;
        if (method === "session/update" && update?.sessionUpdate === "available_commands_update") {
          supportsUsageCommand = update.availableCommands?.some(command => command.name.replace(/^\//, "") === "usage") ?? false;
          if (session) session.supportsUsageCommand = supportsUsageCommand;
        }
        if (session) this.onNotification(session, method, params);
      },
      onRequest: (method, params) => this.onServerRequest(method, params),
    });

    try {
      await peer.request("initialize", {
        protocolVersion: 1,
        clientCapabilities: {},
        clientInfo: { name: "excel-ai-addin", title: "Excel AI Add-in", version: "0.1.0" },
      });
      const created = await peer.request<{ sessionId: string; configOptions?: EffortConfigOption[] }>(
        "session/new",
        {
          cwd,
          mcpServers: [],
        },
        30_000,
      );
      const effort = effortConfig(created.configOptions, ctx.effort ?? "medium");
      if (effort) await peer.request("session/set_config_option", { sessionId: created.sessionId, ...effort });
      session = { peer, configDir, acpSessionId: created.sessionId, primed: false, active: null, supportsUsageCommand };
      this.sessions.set(providerSessionId, session);
      peer.onExit(() => {
        const current = this.sessions.get(providerSessionId);
        current?.active?.queue.fail(
          new Error(`copilot exited unexpectedly: ${peer.stderrTail.slice(-3).join(" | ")}`),
        );
        this.sessions.delete(providerSessionId);
        fs.rmSync(configDir, { recursive: true, force: true });
      });
      return providerSessionId;
    } catch (error) {
      await peer.close();
      fs.rmSync(configDir, { recursive: true, force: true });
      throw error;
    }
  }

  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const { requestId } = input;
    if (!session || !session.peer.alive) {
      yield { type: "error", requestId, code: "no_session", message: "copilot session is not running" };
      yield { type: "turn_finished", requestId, reason: "error" };
      return;
    }

    const queue = new AsyncQueue<ProviderEvent>();
    const active: ActiveTurn = { queue, requestId, text: "", tools: new Map(), finishedTools: new Set() };
    session.active = active;
    yield { type: "turn_started", requestId };

    const cancel = () => session.peer.notify("session/cancel", { sessionId: session.acpSessionId });
    input.signal.addEventListener("abort", cancel, { once: true });

    // ACP has no system-prompt field, so the instructions travel with the first prompt.
    const userText = buildTurnPrompt(input.prompt, input.selection);
    const text = session.primed ? userText : `${SYSTEM_PROMPT}\n\n---\n\n${userText}`;
    session.primed = true;

    session.peer
      .request<{ stopReason?: string; usage?: { inputTokens: number; outputTokens: number; totalTokens: number; cachedReadTokens?: number } }>(
        "session/prompt",
        { sessionId: session.acpSessionId, prompt: [{ type: "text", text }] },
        30 * 60_000,
      )
      .then(async (response) => {
        this.flushText(active);
        // Copilot's ACP PromptResponse usage contains CLI-session totals.
        // Convert to the increment for this turn before the bridge adds it.
        // usage_update.used is context occupancy and must never be counted here.
        if (response.usage) {
          const usage = TokenUsageSchema.safeParse({ ...response.usage, cachedInputTokens: response.usage.cachedReadTokens ?? 0 });
          if (usage.success) {
            session.usage ??= new CumulativeUsage();
            const increment = session.usage.next(usage.data);
            if (session.supportsUsageCommand && !input.signal.aborted) {
              try {
                // Advertised informational slash command: no model call or credits consumed.
                // Suppress its messages so the usage report never appears as an assistant reply.
                session.usageReport = "";
                await session.peer.request("session/prompt", { sessionId: session.acpSessionId,
                  prompt: [{ type: "text", text: "/usage" }] }, 5000);
                const total = parseCopilotCredits(session.usageReport);
                if (total !== null) {
                  increment.copilotCredits = Math.max(0, total - (session.creditTotal ?? 0));
                  session.creditTotal = total;
                }
              } catch (error) {
                this.runtime.logger.debug("Copilot credit report unavailable", { error: String(error) });
              } finally { session.usageReport = undefined; }
            }
            queue.push({ type: "usage_update", requestId, usage: increment });
          }
        }
        const stop = response?.stopReason ?? "end_turn";
        if (stop === "cancelled") {
          queue.push({ type: "turn_finished", requestId, reason: "cancelled" });
        } else if (stop === "end_turn" || stop === "max_turn_requests") {
          queue.push({ type: "turn_finished", requestId, reason: "completed" });
        } else {
          queue.push({
            type: "error",
            requestId,
            code: "provider_error",
            message: `copilot stopped the turn (${stop})`,
          });
          queue.push({ type: "turn_finished", requestId, reason: "error" });
        }
        queue.end();
      })
      .catch((error) => {
        this.flushText(active);
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
      input.signal.removeEventListener("abort", cancel);
      if (session.active === active) session.active = null;
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
    await Promise.all([...this.sessions.keys()].map((id) => this.closeSession(id)));
  }

  private flushText(active: ActiveTurn): void {
    if (!active.text) return;
    active.queue.push({ type: "message_completed", requestId: active.requestId, text: active.text });
    active.text = "";
  }

  private onNotification(session: CopilotSession, method: string, raw: unknown): void {
    if (method !== "session/update") return;
    const params = (raw ?? {}) as { sessionId?: string; update?: Record<string, any> };
    const active = session.active;
    const update = params.update;
    if (session.usageReport !== undefined) {
      if (update?.sessionUpdate === "agent_message_chunk" && update.content?.type === "text") {
        session.usageReport += String(update.content.text ?? "");
      }
      return;
    }
    if (!active || !update || params.sessionId !== session.acpSessionId) return;
    const { queue, requestId } = active;
    if (update.sessionUpdate !== "agent_thought_chunk") {
      this.runtime.logger.debug("copilot update", update);
    }

    switch (update.sessionUpdate) {
      case "agent_message_chunk": {
        const text = update.content?.type === "text" ? String(update.content.text ?? "") : "";
        // The CLI announces its own tool filtering as a chat chunk; it is not model output.
        if (CLI_NOTICE.test(text)) break;
        if (text) {
          active.text += text;
          queue.push({ type: "message_delta", requestId, text });
        }
        break;
      }
      case "tool_call": {
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

  private onServerRequest(method: string, params: unknown): unknown {
    if (method === "session/request_permission") {
      const request = params as {
        toolCall?: Record<string, any>;
        options?: Array<{ optionId: string; kind?: string }>;
      };
      const options = request.options ?? [];
      const isExcel = toolNameOf(request.toolCall ?? {}, true).isExcel;
      const wanted = isExcel ? ["allow_once", "allow_always"] : ["reject_once", "reject_always"];
      const choice = options.find((o) => wanted.includes(o.kind ?? ""));
      if (choice) return { outcome: { outcome: "selected", optionId: choice.optionId } };
      return { outcome: { outcome: "cancelled" } };
    }
    throw new Error(`unsupported client method: ${method}`);
  }
}

/** Copilot CLI status line emitted as an `agent_message_chunk` when `--available-tools` is used. */
const CLI_NOTICE = /^Info: Disabled tools: /;

function isAuthError(error: unknown): boolean {
  return /auth|login|401|403|token/i.test(error instanceof Error ? error.message : String(error));
}

function contentText(update: Record<string, any>): string {
  const parts: string[] = [];
  for (const item of update.content ?? []) {
    const text = item?.content?.text ?? item?.text;
    if (typeof text === "string") parts.push(text);
  }
  if (parts.length === 0 && update.rawOutput) {
    parts.push(typeof update.rawOutput === "string" ? update.rawOutput : JSON.stringify(update.rawOutput));
  }
  return parts.join(" ");
}

/** Extract the bare Excel tool name from an ACP tool_call (`excel-get_selection`, `excel/get_selection`, ...). */
function toolNameOf(update: Record<string, any>): string;
function toolNameOf(update: Record<string, any>, detail: true): { name: string; isExcel: boolean };
function toolNameOf(update: Record<string, any>, detail?: true): string | { name: string; isExcel: boolean } {
  const candidates = [update._meta?.toolName, update.rawInput?.toolName, update.title, update.kind]
    .filter((v): v is string => typeof v === "string");
  let name = candidates[0] ?? "unknown";
  let isExcel = false;
  for (const candidate of candidates) {
    const match = new RegExp(`${MCP_SERVER_NAME}[-_/:.]+([a-z_]+)`, "i").exec(candidate);
    if (match) {
      name = match[1]!;
      isExcel = true;
      break;
    }
  }
  return detail ? { name, isExcel } : name;
}
