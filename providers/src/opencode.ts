import { randomUUID } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderEvent,
  ProviderId,
} from "../../protocol/src/index.js";
import { acpPermissionOutcome, contentText, toolNameOf } from "./acp.js";
import { AsyncQueue } from "./async-queue.js";
import { BaseCliAdapter, cliEnvironment, execCli, type CliProbe } from "./base.js";
import { effortConfig, type EffortConfigOption } from "./effort.js";
import { JsonRpcPeer } from "./jsonrpc.js";
import { buildTurnPrompt, SYSTEM_PROMPT } from "./prompt.js";
import { CumulativeUsage } from "./usage.js";
import { MCP_SERVER_NAME, type ProviderModelCatalog, type SendTurnInput, type StartSessionContext } from "./types.js";
import { ensureWorkspaceDir } from "./workspace.js";
import { nativeTitle, validNativeId } from "./session-metadata.js";
import { parseOpencodeExport } from "./opencode-usage.js";

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
 * rejects everything else when the CLI asks for permission.
 */
export class OpencodeAdapter extends BaseCliAdapter {
  private titleCache?: { at: number; value: Promise<Map<string, string>> };
  async sessionTitle(id: string): Promise<string | undefined> {
    if (!validNativeId(id)) return undefined;
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

  private async originalUsage(id: string): Promise<ReturnType<typeof parseOpencodeExport>> {
    const result = await execCli(this.command, ["export", id, "--pure"], this.env, 10_000);
    if (result.code !== 0) throw new Error("OpenCode original session export is unavailable");
    return parseOpencodeExport(JSON.parse(result.stdout), id);
  }
  readonly id: ProviderId = "opencode";
  readonly capabilities: ProviderCapabilities = {
    streaming: true,
    cancel: true,
    sessionResume: true,
    mcp: true,
    acp: true,
  };

  /** Lowest OpenCode version this adapter has been tested with. */
  static readonly MIN_VERSION = "1.0.0";

  protected override get minVersion(): string {
    return OpencodeAdapter.MIN_VERSION;
  }

  private readonly sessions = new Map<string, OpencodeSession>();

  private get command(): string {
    return process.env.OPENCODE_BIN ?? "opencode";
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
    const { dir: configDir, path: configPath } = this.writePermissionConfig();

    let session: OpencodeSession | null = null;
    const peer = new JsonRpcPeer({
      command: this.command,
      args: ["acp"],
      cwd,
      // The permission config is merged on top of the user's own OpenCode
      // config, so their providers, agents and skills keep working.
      env: { ...this.env, OPENCODE_CONFIG: configPath },
      versionField: true,
      onStderr: (line) => log.debug(`opencode stderr: ${line}`),
      onNotification: (method, params) => session && this.onNotification(session, method, params),
      onRequest: (method, params) => this.onServerRequest(method, params),
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
        peer,
        acpSessionId: created.sessionId,
        configOptions: created.configOptions ?? [],
        model: null,
        primed: Boolean(ctx.resumeSessionId),
        active: null,
        configDir,
      };
      if (ctx.resumeSessionId) {
        session.usage = new CumulativeUsage();
        try { session.usage.next((await this.originalUsage(created.sessionId)).usage); }
        catch { session.usageBaselinePending = true; }
      }
      this.sessions.set(providerSessionId, session);
      peer.onExit(() => {
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
      await peer.close();
      fs.rmSync(configDir, { recursive: true, force: true });
      throw error;
    }
  }

  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const { requestId } = input;
    if (!session || !session.peer.alive) {
      yield { type: "error", requestId, code: "no_session", message: "opencode session is not running" };
      yield { type: "turn_finished", requestId, reason: "error" };
      return;
    }

    const queue = new AsyncQueue<ProviderEvent>();
    const active: ActiveTurn = { queue, requestId, text: "", tools: new Map(), finishedTools: new Set() };
    session.active = active;
    yield { type: "session_info", requestId, provider: this.id, providerSessionId: session.acpSessionId,
      title: session.title };
    yield { type: "turn_started", requestId };

    const cancel = () => session.peer.notify("session/cancel", { sessionId: session.acpSessionId });
    input.signal.addEventListener("abort", cancel, { once: true });

    // ACP has no system-prompt field, so the instructions travel with the first prompt.
    const userText = buildTurnPrompt(input.prompt, input.selection);
    const text = session.primed ? userText : `${input.systemPrompt ?? SYSTEM_PROMPT}\n\n---\n\n${userText}`;
    session.primed = true;

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
      return session.peer.request<{
        stopReason?: string;
        usage?: { inputTokens: number; outputTokens: number; totalTokens: number; cachedReadTokens?: number };
      }>(
        "session/prompt",
        { sessionId: session.acpSessionId, prompt: [{ type: "text", text }] },
        30 * 60_000,
      );
    };

    run()
      .then(async (response) => {
        this.flushText(active);
        await this.reportUsage(session, active);
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
            message: `opencode stopped the turn (${stop})`,
          });
          queue.push({ type: "turn_finished", requestId, reason: "error" });
        }
        queue.end();
      })
      .catch(async (error) => {
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
      input.signal.removeEventListener("abort", cancel);
      if (session.active === active) session.active = null;
    }
  }

  private async reportUsage(session: OpencodeSession, active: ActiveTurn): Promise<void> {
    try {
      const snapshot = await this.originalUsage(session.acpSessionId);
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
    await Promise.all([...this.sessions.keys()].map((id) => this.closeSession(id)));
  }

  /**
   * Permission config handed to the CLI: every tool has to be approved, and
   * only the Excel MCP tools are pre-approved. Everything else reaches this
   * adapter as a permission request and is rejected.
   */
  private writePermissionConfig(): { dir: string; path: string } {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-ai-opencode-"));
    const configPath = path.join(dir, "opencode.json");
    fs.writeFileSync(
      configPath,
      JSON.stringify({
        $schema: "https://opencode.ai/config.json",
        permission: {
          "*": "ask",
          [`${MCP_SERVER_NAME}_*`]: "allow",
        },
      }),
      { mode: 0o600 },
    );
    return { dir, path: configPath };
  }

  private async selectModel(session: OpencodeSession, model: string): Promise<void> {
    try {
      const result = await session.peer.request<{ configOptions?: OpencodeConfigOption[] }>(
        "session/set_config_option",
        { sessionId: session.acpSessionId, configId: "model", value: model },
        10_000,
      );
      if (result?.configOptions) session.configOptions = result.configOptions;
      session.model = model;
    } catch (error) {
      this.runtime.logger.warn("opencode could not select the model", { model, error: String(error) });
    }
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
    if (update.sessionUpdate === "config_option_update") {
      if (Array.isArray(update.configOptions)) session.configOptions = update.configOptions;
      return;
    }
    const active = session.active;
    if (!active || params.sessionId !== session.acpSessionId) return;
    const { queue, requestId } = active;

    switch (update.sessionUpdate) {
      case "agent_message_chunk": {
        const text = update.content?.type === "text" ? String(update.content.text ?? "") : "";
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
    if (method === "session/request_permission") return acpPermissionOutcome(params);
    throw new Error(`unsupported client method: ${method}`);
  }
}

function isAuthError(error: unknown): boolean {
  return /auth|login|401|403|token|credential/i.test(error instanceof Error ? error.message : String(error));
}

function stripAnsi(value: string): string {
  return value.replace(/\u001b\[[0-9;]*m/g, "");
}
