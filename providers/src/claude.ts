import type { ChildProcess } from "node:child_process";
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

interface ClaudeSession {
  mcp: McpEndpoint;
  /** Path of the 0600 MCP config file holding the session token. */
  mcpConfigPath: string;
  /** Conversation id; becomes resumable after the first successful turn. */
  conversationId: string;
  resumable: boolean;
  child: ChildProcess | null;
}

const MCP_PREFIX = `mcp__${MCP_SERVER_NAME}__`;

/**
 * Claude Code CLI.
 *
 * One `claude -p --output-format stream-json` process per turn, continued
 * across turns with `--resume <session-id>`. The Excel MCP server is injected
 * with `--mcp-config` (+ `--strict-mcp-config`), built-in tools are disabled
 * and only the Excel MCP tools are pre-approved. Cancel = SIGTERM to the
 * running turn process; the conversation stays resumable.
 */
export class ClaudeAdapter extends BaseCliAdapter {
  async listModels(): Promise<ProviderModelCatalog> {
    return {
      models: [
        { id: "sonnet", name: "Sonnet (latest)" },
        { id: "opus", name: "Opus (latest)" },
        { id: "haiku", name: "Haiku (latest)" },
      ],
      defaultModelId: null,
    };
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
    const providerSessionId = `claude-${randomUUID()}`;
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-ai-claude-"));
    const mcpConfigPath = path.join(dir, "mcp.json");
    fs.writeFileSync(
      mcpConfigPath,
      JSON.stringify({
        mcpServers: {
          [MCP_SERVER_NAME]: { command: ctx.mcp.command, args: ctx.mcp.args, env: ctx.mcp.env },
        },
      }),
      { mode: 0o600 },
    );
    this.sessions.set(providerSessionId, {
      mcp: ctx.mcp,
      mcpConfigPath,
      conversationId: randomUUID(),
      resumable: false,
      child: null,
    });
    return providerSessionId;
  }

  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    const session = this.sessions.get(input.providerSessionId);
    const { requestId } = input;
    if (!session) {
      yield { type: "error", requestId, code: "no_session", message: "claude session is not running" };
      yield { type: "turn_finished", requestId, reason: "error" };
      return;
    }
    yield { type: "turn_started", requestId };
    if (input.signal.aborted) {
      yield { type: "turn_finished", requestId, reason: "cancelled" };
      return;
    }

    const model = input.model ?? process.env.CLAUDE_MODEL;
    const args = [
      "-p",
      "--output-format",
      "stream-json",
      "--verbose",
      "--include-partial-messages",
      ...(session.resumable ? ["--resume", session.conversationId] : ["--session-id", session.conversationId]),
      "--mcp-config",
      session.mcpConfigPath,
      "--strict-mcp-config",
      "--tools",
      "",
      "--allowedTools",
      `mcp__${MCP_SERVER_NAME}`,
      "--permission-mode",
      "dontAsk",
      "--system-prompt",
      SYSTEM_PROMPT,
      "--disable-slash-commands",
      // The scratch cwd is empty, so this skips the user's own hooks, CLAUDE.md
      // and plugins (auth is unaffected) and keeps the Excel session predictable.
      "--setting-sources",
      "project",
      ...(model ? ["--model", model] : []),
      "--effort",
      input.effort ?? "medium",
    ];

    const child = spawn(this.command, args, {
      cwd: ensureWorkspaceDir(),
      env: { ...this.env, CLAUDE_CODE_EFFORT_LEVEL: input.effort ?? "medium" },
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
    });
    session.child = child;
    child.stdin!.on("error", () => undefined);
    // The prompt (which may contain workbook data) goes over stdin, never argv.
    child.stdin!.end(buildTurnPrompt(input.prompt, input.selection));

    const stderrTail: string[] = [];
    createInterface({ input: child.stderr! }).on("line", (line) => {
      stderrTail.push(line);
      if (stderrTail.length > 40) stderrTail.shift();
      this.runtime.logger.debug(`claude stderr: ${line}`);
    });

    const onAbort = () => terminate(child);
    input.signal.addEventListener("abort", onAbort, { once: true });

    const exited = new Promise<{ code: number | null; signal: NodeJS.Signals | null }>((resolve) => {
      child.once("exit", (code, signal) => resolve({ code, signal }));
      child.once("error", () => resolve({ code: null, signal: null }));
    });

    const translator = new ClaudeStreamTranslator(requestId);
    let finished = false;
    try {
      for await (const line of createInterface({ input: child.stdout!, crlfDelay: Infinity })) {
        const trimmed = line.trim();
        if (!trimmed.startsWith("{")) continue;
        let message: Record<string, any>;
        try {
          message = JSON.parse(trimmed);
        } catch {
          continue;
        }
        if (message.type === "result" && !message.is_error && message.session_id) {
          session.conversationId = String(message.session_id);
          session.resumable = true;
        }
        for (const event of translator.translate(message)) {
          if (event.type === "turn_finished") {
            finished = true;
            yield input.signal.aborted ? { ...event, reason: "cancelled" } : event;
          } else {
            yield event;
          }
        }
      }

      const { code, signal } = await exited;
      if (!finished) {
        if (input.signal.aborted) {
          yield { type: "turn_finished", requestId, reason: "cancelled" };
        } else {
          const detail = stderrTail.slice(-3).join(" | ");
          yield {
            type: "error",
            requestId,
            code: "provider_error",
            message: `claude exited without a result (code=${code}, signal=${signal})${detail ? `: ${detail}` : ""}`,
          };
          yield { type: "turn_finished", requestId, reason: "error" };
        }
      }
    } finally {
      input.signal.removeEventListener("abort", onAbort);
      // A turn that never completed may have left a half-created conversation;
      // start the next attempt on a fresh id instead of colliding with it.
      if (!session.resumable) session.conversationId = randomUUID();
      if (child.exitCode === null && child.signalCode === null) terminate(child);
      if (session.child === child) session.child = null;
    }
  }

  override async cancelTurn(providerSessionId: string): Promise<void> {
    const child = this.sessions.get(providerSessionId)?.child;
    if (child) terminate(child);
  }

  override async closeSession(providerSessionId: string): Promise<void> {
    const session = this.sessions.get(providerSessionId);
    this.sessions.delete(providerSessionId);
    if (!session) return;
    if (session.child) terminate(session.child);
    fs.rmSync(path.dirname(session.mcpConfigPath), { recursive: true, force: true });
  }

  override async dispose(): Promise<void> {
    await Promise.all([...this.sessions.keys()].map((id) => this.closeSession(id)));
  }
}

function terminate(child: ChildProcess): void {
  if (child.exitCode !== null || child.signalCode !== null) return;
  child.kill("SIGTERM");
  const timer = setTimeout(() => child.kill("SIGKILL"), 3000);
  timer.unref();
  child.once("exit", () => clearTimeout(timer));
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
          events.push({ type: "usage_update", requestId, usage: { inputTokens, outputTokens, cachedInputTokens, totalTokens: inputTokens + outputTokens } });
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
