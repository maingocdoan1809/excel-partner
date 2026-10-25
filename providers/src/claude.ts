import spawn from "cross-spawn";
import { randomUUID } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createInterface } from "node:readline";
import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderEvent,
  ProviderId,
} from "../../protocol/src/index.js";
import { BaseCliAdapter, cliEnvironment, execCli, type CliProbe } from "./base.js";
import { buildTurnPrompt, SYSTEM_PROMPT } from "./prompt.js";
import { MCP_SERVER_NAME, type McpEndpoint, type ProviderModelCatalog, type SendTurnInput, type StartSessionContext } from "./types.js";
import { ensureWorkspaceDir } from "./workspace.js";
import { claudeTitle, validNativeId } from "./session-metadata.js";
import { answerClaudeQuestion, type RequestUserInput } from "./user-input.js";
import { cachedCommands, normalizeCommands } from "./commands.js";
import { discoverClaudeCommands } from "./command-discovery.js";
import { ClaudeProcess } from "./claude-process.js";
import { AsyncQueue } from "./async-queue.js";

interface ClaudeSession {
  requestUserInput?: RequestUserInput;
  mcp: McpEndpoint;
  /** Path of the 0600 MCP config file holding the session token. */
  mcpConfigPath: string;
  /** Conversation id; becomes resumable when the CLI confirms initialization. */
  conversationId: string;
  resumable: boolean;
  process: ClaudeProcess;
  active: ClaudeTurn | null;
}

interface ClaudeTurn {
  input: SendTurnInput;
  queue: AsyncQueue<ProviderEvent>;
  translator: ClaudeStreamTranslator;
  questions: Map<string, AbortController>;
  interruptSent: boolean;
  interruptController: AbortController;
}

/** Shape of one entry in the CLI's `initialize` control response (`models`). */
interface ClaudeModelInfo { value?: unknown; displayName?: unknown; description?: unknown; resolvedModel?: unknown }

/** Turn the CLI's own model list (the one behind `/model`) into the picker catalog. */
export function parseClaudeModels(models: unknown): ProviderModelCatalog | null {
  if (!Array.isArray(models)) return null;
  const seen = new Set<string>();
  const out: ProviderModelCatalog["models"] = [];
  let defaultModelId: string | null = null;
  for (const item of models as ClaudeModelInfo[]) {
    if (typeof item?.value !== "string" || !item.value || seen.has(item.value)) continue;
    seen.add(item.value);
    if (item.value === "default") {
      if (typeof item.resolvedModel === "string") defaultModelId = item.resolvedModel;
      continue; // the picker already has its own "Default" entry
    }
    const display = typeof item.displayName === "string" && item.displayName ? item.displayName : item.value;
    // Some accounts report only "Sonnet" as displayName and keep the version in the
    // description ("Sonnet 5.5 · Efficient for routine tasks"). Prefer that lead-in.
    const lead = typeof item.description === "string" ? item.description.split(" · ")[0]!.trim() : "";
    out.push({ id: item.value, name: lead.toLowerCase().startsWith(display.toLowerCase()) && lead.length <= 40 ? lead : display });
  }
  return out.length ? { models: out, defaultModelId } : null;
}

const MCP_PREFIX = `mcp__${MCP_SERVER_NAME}__`;

/**
 * Claude Code CLI.
 *
 * One persistent stream-json process per session. Each result completes one
 * turn; stdin stays open for the next prompt. Resume is only used when opening
 * an existing conversation. Interrupt cancels a turn without closing its runtime.
 */
export class ClaudeAdapter extends BaseCliAdapter {
  private readonly commandCatalog = cachedCommands(() => discoverClaudeCommands(this.command, this.env, ensureWorkspaceDir()));
  async listCommands() {
    const commands = await this.commandCatalog();
    return commands.some(command => command.name === "skills") ? commands : [...commands, ...normalizeCommands([
      { name: "skills", description: "Browse Claude commands and skills", inputHint: "list | <name> <request>" },
    ])];
  }
  async sessionTitle(id: string): Promise<string | undefined> { return claudeTitle(id, ensureWorkspaceDir(), this.env); }
  /**
   * Ask the CLI itself: the `initialize` control request of the stream-json
   * protocol returns the same account-filtered model list as `/model`
   * (exact versions such as "Sonnet 5" and "Sonnet 5.5"). No model call is made.
   * Falls back to the three basic aliases if the CLI cannot answer.
   */
  async listModels(): Promise<ProviderModelCatalog> {
    const fallback: ProviderModelCatalog = {
      models: [
        { id: "sonnet", name: "Sonnet (latest)" },
        { id: "opus", name: "Opus (latest)" },
        { id: "haiku", name: "Haiku (latest)" },
      ],
      defaultModelId: null,
    };
    try { return (await this.queryModels()) ?? fallback; } catch { return fallback; }
  }

  private queryModels(timeoutMs = 20_000): Promise<ProviderModelCatalog | null> {
    return new Promise((resolve) => {
      const child = spawn(this.command, [
        "-p", "--input-format", "stream-json", "--output-format", "stream-json", "--verbose",
        "--setting-sources", "project", "--tools", "", "--disable-slash-commands",
      ], { cwd: ensureWorkspaceDir(), env: this.env, windowsHide: true, stdio: ["pipe", "pipe", "ignore"] });
      let done = false;
      const finish = (value: ProviderModelCatalog | null) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        child.kill();
        resolve(value);
      };
      const timer = setTimeout(() => finish(null), timeoutMs);
      child.on("error", () => finish(null));
      child.on("exit", () => finish(null));
      createInterface({ input: child.stdout! }).on("line", (line) => {
        try {
          const message = JSON.parse(line) as { type?: string; response?: { request_id?: string; response?: { models?: unknown } } };
          if (message.type === "control_response" && message.response?.request_id === "models") finish(parseClaudeModels(message.response.response?.models));
        } catch { /* not JSON */ }
      });
      child.stdin!.on("error", () => finish(null));
      child.stdin!.write(JSON.stringify({ type: "control_request", request_id: "models", request: { subtype: "initialize" } }) + "\n");
    });
  }
  readonly id: ProviderId = "claude";
  readonly capabilities: ProviderCapabilities = {
    streaming: true,
    cancel: true,
    sessionResume: true,
    mcp: true,
  };

  /** Lowest Claude Code version this adapter has been tested with. */
  static readonly MIN_VERSION = "2.1.294";

  protected override get minVersion(): string {
    return ClaudeAdapter.MIN_VERSION;
  }

  private readonly sessions = new Map<string, ClaudeSession>();

  protected probe(): CliProbe {
    return {
      command: this.command,
      versionArgs: ["--version"],
      installHint: "npm install -g @anthropic-ai/claude-code",
      loginHint: "claude",
    };
  }

  private get command(): string {
    return process.env.CLAUDE_BIN ?? "claude";
  }

  protected override async checkLogin(
    command: string,
  ): Promise<{ status: ProviderConnectionStatus; hint: string | null }> {
    const result = await execCli(command, ["auth", "status"], this.env);
    try {
      const parsed = JSON.parse(result.stdout) as { loggedIn?: boolean };
      if (parsed.loggedIn) return { status: "ready", hint: null };
    } catch {
      /* fall through */
    }
    if (process.env.ANTHROPIC_API_KEY) return { status: "ready", hint: null };
    return { status: "login_required", hint: `Run \`claude\` on ${cliEnvironment()} and complete /login, then retry.` };
  }

  override async startSession(ctx: StartSessionContext): Promise<string> {
    if (ctx.resumeSessionId && !validNativeId(ctx.resumeSessionId)) throw new Error("Invalid Claude resume session ID");
    ctx.signal.throwIfAborted();
    const providerSessionId = `claude-${randomUUID()}`;
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-ai-claude-"));
    const mcpConfigPath = path.join(dir, "mcp.json");
    let transport: ClaudeProcess | undefined;
    try {
      fs.writeFileSync(mcpConfigPath, JSON.stringify({ mcpServers: {
        [MCP_SERVER_NAME]: { command: ctx.mcp.command, args: ctx.mcp.args, env: ctx.mcp.env },
      } }), { mode: 0o600 });
      const conversationId = ctx.resumeSessionId ?? randomUUID();
      const model = ctx.model ?? process.env.CLAUDE_MODEL;
      const systemPrompt = ctx.systemPrompt ?? SYSTEM_PROMPT;
      const webTools = ctx.mcp.env.EXCEL_WEB_SEARCH === "1" ? ["WebSearch", "WebFetch"] : [];
      let session: ClaudeSession | undefined;
      transport = new ClaudeProcess({
        command: this.command,
        args: [
          "-p", "--input-format", "stream-json", "--permission-prompt-tool", "stdio",
          "--output-format", "stream-json", "--verbose", "--include-partial-messages",
          ...(ctx.resumeSessionId ? ["--resume", conversationId] : ["--session-id", conversationId]),
          "--mcp-config", mcpConfigPath, "--strict-mcp-config",
          "--tools", [...(ctx.mcp.toolNames ? [] : ["AskUserQuestion"]), ...webTools].join(","),
          "--allowedTools", `mcp__${MCP_SERVER_NAME}`, ...webTools,
          "--permission-mode", "default",
          ...(systemPrompt.trim() ? ["--system-prompt", systemPrompt] : []),
          // Preserve the existing restriction to the empty scratch workspace.
          "--setting-sources", "project",
          ...(model ? ["--model", model] : []), "--effort", ctx.effort ?? "medium",
        ],
        cwd: ensureWorkspaceDir(),
        env: { ...this.env, CLAUDE_CODE_EFFORT_LEVEL: ctx.effort ?? "medium" },
        onStderr: line => this.runtime.logger.debug(`claude stderr: ${line}`),
        onMessage: message => { if (session) this.onMessage(session, message); },
        onRuntimeClosed: ctx.onRuntimeClosed,
        onExit: error => {
          const active = session?.active;
          if (!active) return;
          if (!active.input.signal.aborted) active.queue.push({ type: "error", requestId: active.input.requestId,
            code: "provider_error", message: error.message });
          this.finish(session!, active.input.signal.aborted ? "cancelled" : "error");
        },
      });
      session = { requestUserInput: ctx.requestUserInput, mcp: ctx.mcp, mcpConfigPath,
        conversationId, resumable: Boolean(ctx.resumeSessionId), process: transport, active: null };
      await transport.request("initialize", ctx.signal);
      ctx.signal.throwIfAborted();
      this.sessions.set(providerSessionId, session);
      return providerSessionId;
    } catch (error) {
      await transport?.close();
      fs.rmSync(dir, { recursive: true, force: true });
      throw error;
    }
  }

  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const { requestId } = input;
    if (!session || !session.process.alive) throw new Error("Claude session is not running; reconnect to resume the conversation");
    if (session.active) throw new Error("Another Claude turn is already running");
    if (input.signal.aborted) { yield { type: "turn_finished", requestId, reason: "cancelled" }; return; }
    if (session.resumable) yield { type: "session_info", requestId, provider: this.id,
      providerSessionId: session.conversationId, title: await this.sessionTitle(session.conversationId) };
    yield { type: "turn_started", requestId };

    if (input.command?.name === "skills") {
      const commands = await this.listCommands();
      const args = input.command.args;
      if (!args || args === "list") {
        yield { type: "message_completed", requestId, text: "Available Claude commands and skills:\n\n" + commands.filter(command => command.name !== "skills")
          .map(command => `- **/${command.name}** — ${command.description}`).join("\n") + "\n\nRun a skill with `/<name> <request>`." };
        yield { type: "turn_finished", requestId, reason: "completed" };
        return;
      }
      const [, name, task] = /^(\S+)(?:\s+([\s\S]*))?$/.exec(args)!;
      const target = commands.find(command => command.name === name && command.name !== "skills" && command.execution === "native");
      if (!target) throw new Error(`Unknown Claude skill or command: ${name}`);
      input = { ...input, command: { name: target.name, args: task ?? "" } };
    }
    const active: ClaudeTurn = { input, queue: new AsyncQueue(), translator: new ClaudeStreamTranslator(requestId),
      questions: new Map(), interruptSent: false, interruptController: new AbortController() };
    session.active = active;
    const abort = () => { void this.interrupt(session); };
    input.signal.addEventListener("abort", abort, { once: true });
    try {
      input.signal.throwIfAborted();
      const text = input.command ? `/${input.command.name}${input.command.args ? ` ${input.command.args}` : ""}` : buildTurnPrompt(input.prompt, input.selection);
      const content = input.images?.length ? [{ type: "text", text }, ...input.images.map(image => ({ type: "image", source: { type: "base64", media_type: image.mediaType, data: image.data } }))] : text;
      session.process.write({ type: "user", session_id: session.conversationId, parent_tool_use_id: null,
        message: { role: "user", content } });
      for await (const event of active.queue) {
        // Read the native title at completion, outside the stdout callback.
        if (event.type === "session_info") yield { ...event, title: await this.sessionTitle(event.providerSessionId) };
        else yield event;
      }
    } finally {
      active.interruptController.abort();
      input.signal.removeEventListener("abort", abort);
      for (const question of active.questions.values()) question.abort();
      if (session.active === active) {
        // An abandoned consumer must not leave a live turn feeding a future one.
        await session.process.close();
        session.active = null;
      }
    }
  }

  private onMessage(session: ClaudeSession, message: Record<string, any>): void {
    const active = session.active;
    if (!active) return; // Late notifications belong to the completed turn.
    const { requestId } = active.input;
    if (message.type === "control_cancel_request") {
      active.questions.get(message.request_id)?.abort();
      return;
    }
    if (message.type === "control_request") {
      void this.answerControl(session, active, message);
      return;
    }
    const confirmsSession = (message.type === "result" || (message.type === "system" && message.subtype === "init"))
      && typeof message.session_id === "string" && validNativeId(message.session_id);
    if (confirmsSession) {
      if (session.resumable && session.conversationId !== message.session_id) {
        active.queue.push({ type: "error", requestId, code: "provider_error", message: "Claude returned a different session while resuming" });
        this.finish(session, "error");
        void session.process.close();
        return;
      }
      session.conversationId = message.session_id;
      session.resumable = true;
    }
    if (confirmsSession || (message.type === "result" && session.resumable)) active.queue.push({ type: "session_info", requestId,
      provider: this.id, providerSessionId: session.conversationId });
    for (const event of active.translator.translate(message)) {
      if (event.type === "turn_finished") this.finish(session, active.input.signal.aborted ? "cancelled" : event.reason);
      else active.queue.push(event.type === "usage_update" ? { ...event,
        usage: { ...event.usage, source: "claude_result", providerSessionId: session.conversationId } } : event);
    }
  }

  private async answerControl(session: ClaudeSession, active: ClaudeTurn, message: Record<string, any>): Promise<void> {
    const controller = new AbortController();
    active.questions.set(message.request_id, controller);
    let response: unknown = { behavior: "deny", message: "This tool is unavailable in the add-in." };
    try {
      if (message.request?.subtype === "can_use_tool" && message.request.tool_name === "AskUserQuestion") {
        response = await answerClaudeQuestion(message.request.input, session.requestUserInput, controller.signal);
      }
      if (session.active === active && !controller.signal.aborted && !active.input.signal.aborted) {
        session.process.write({ type: "control_response", response: { subtype: "success", request_id: message.request_id, response } });
      }
    } catch (error) {
      this.runtime.logger.warn("Claude permission response failed", { error: String(error) });
      if (session.active === active) await session.process.close();
    } finally { active.questions.delete(message.request_id); }
  }

  private finish(session: ClaudeSession, reason: "completed" | "cancelled" | "error"): void {
    const active = session.active;
    if (!active) return;
    session.active = null;
    active.interruptController.abort();
    for (const question of active.questions.values()) question.abort();
    active.queue.push({ type: "turn_finished", requestId: active.input.requestId, reason });
    active.queue.end();
  }

  private async interrupt(session: ClaudeSession): Promise<void> {
    const active = session.active;
    if (!active || active.interruptSent) return;
    active.interruptSent = true;
    for (const question of active.questions.values()) question.abort();
    try { await session.process.request("interrupt", active.interruptController.signal, 5000); }
    catch {
      // Completion may race an interrupt ACK. Its timeout must never close a
      // process that has already moved on to another turn.
      if (session.active === active) await session.process.close();
    }
  }

  override async cancelTurn(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    if (session) await this.interrupt(session);
  }

  override async closeSession(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    this.sessions.delete(providerSessionId);
    if (!session) return;
    try { await session.process.close(); }
    finally { fs.rmSync(path.dirname(session.mcpConfigPath), { recursive: true, force: true }); }
  }

  override async dispose(): Promise<void> {
    await Promise.all([...this.sessions.keys()].map(id => this.closeSession(id)));
  }
}

/**
 * Turns Claude Code `stream-json` records into normalised provider events.
 * Pure (no I/O) so it is unit-testable against recorded streams.
 */
export class ClaudeStreamTranslator {
  /** Message ids whose text already streamed as deltas. */
  private readonly streamed = new Set<string>();
  private currentMessageId: string | null = null;
  private readonly toolNames = new Map<string, string>();

  constructor(private readonly requestId: string) {}

  translate(message: Record<string, any>): ProviderEvent[] {
    const { requestId } = this;
    // Sub-agent traffic is not shown to the user.
    if (message.parent_tool_use_id) return [];

    switch (message.type) {
      case "stream_event": {
        const event = message.event ?? {};
        if (event.type === "message_start") {
          this.currentMessageId = event.message?.id ?? null;
        } else if (event.type === "content_block_delta" && event.delta?.type === "text_delta" && event.delta.text) {
          if (this.currentMessageId) this.streamed.add(this.currentMessageId);
          return [{ type: "message_delta", requestId, text: String(event.delta.text) }];
        }
        return [];
      }

      case "assistant": {
        const out: ProviderEvent[] = [];
        const id: string | undefined = message.message?.id;
        for (const block of message.message?.content ?? []) {
          if (block.type === "text" && block.text) {
            if (!(id && this.streamed.has(id))) {
              out.push({ type: "message_delta", requestId, text: String(block.text) });
            }
            out.push({ type: "message_completed", requestId, text: String(block.text) });
          } else if (block.type === "tool_use") {
            const tool = stripPrefix(String(block.name));
            this.toolNames.set(String(block.id), tool);
            out.push({ type: "tool_started", requestId, callId: String(block.id), tool, args: block.input ?? {} });
          }
        }
        return out;
      }

      case "user": {
        const out: ProviderEvent[] = [];
        const content = message.message?.content;
        if (!Array.isArray(content)) return out;
        for (const block of content) {
          if (block.type !== "tool_result") continue;
          const callId = String(block.tool_use_id);
          const failed = block.is_error === true;
          out.push({
            type: "tool_finished",
            requestId,
            callId,
            tool: this.toolNames.get(callId) ?? "unknown",
            ok: !failed,
            error: failed ? toolResultText(block.content).slice(0, 500) || "tool call failed" : null,
          });
        }
        return out;
      }

      case "result": {
        const usage = message.usage;
        const events: ProviderEvent[] = [];
        if (usage && Number.isSafeInteger(usage.input_tokens) && Number.isSafeInteger(usage.output_tokens)) {
          const cachedInputTokens = Math.max(0, usage.cache_read_input_tokens ?? 0);
          const inputTokens = Math.max(0, usage.input_tokens) + cachedInputTokens + Math.max(0, usage.cache_creation_input_tokens ?? 0);
          const outputTokens = Math.max(0, usage.output_tokens);
          events.push({ type: "usage_update", requestId, usage: { inputTokens, outputTokens, cachedInputTokens,
            cacheWriteTokens: Math.max(0, usage.cache_creation_input_tokens ?? 0),
            uncachedInputTokens: Math.max(0, usage.input_tokens), totalTokens: inputTokens + outputTokens } });
        }
        if (message.is_error) {
          const text = String(message.result ?? "claude reported an error");
          const auth = message.api_error_status === 401 || message.api_error_status === 403;
          return [
            ...events,
            { type: "error", requestId, code: auth ? "login_required" : "provider_error", message: text },
            { type: "turn_finished", requestId, reason: "error" },
          ];
        }
        return [...events, { type: "turn_finished", requestId, reason: "completed" }];
      }

      default:
        return [];
    }
  }
}

function stripPrefix(name: string): string {
  return name.startsWith(MCP_PREFIX) ? name.slice(MCP_PREFIX.length) : name;
}

function toolResultText(content: unknown): string {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content.map((c) => (typeof c === "string" ? c : (c as { text?: string }).text ?? "")).join(" ");
  }
  return "";
}
