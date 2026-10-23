import { randomUUID } from "node:crypto";
import os from "node:os";
import path from "node:path";
import fs from "node:fs";
import { createRequire } from "node:module";
import type { CopilotClient, CopilotSession, SessionConfig } from "@github/copilot-sdk";
import { exposedToolNames, UserInputRequestSchema, type ProviderEvent } from "../../protocol/src/index.js";
import { AsyncQueue } from "./async-queue.js";
import { askProviderQuestions } from "./user-input.js";
import { CopilotUsageReader } from "./copilot-usage.js";
import { CumulativeUsage } from "./usage.js";
import { copilotTitle, validNativeId } from "./session-metadata.js";
import { buildTurnPrompt } from "./prompt.js";
import { ensureWorkspaceDir } from "./workspace.js";
import type { SendTurnInput, StartSessionContext } from "./types.js";

interface NativeSession {
  client: CopilotClient;
  session: CopilotSession;
  reader: CopilotUsageReader;
  usage: CumulativeUsage;
  baselinePending: boolean;
  active?: { signal: AbortSignal; controller: AbortController; queue: AsyncQueue<ProviderEvent> };
}

/** Copilot's SDK transport exposes ask_user; its ACP transport does not. */
export class CopilotNativeSessions {
  private readonly sessions = new Map<string, NativeSession>();
  has(id: string) { return this.sessions.has(id); }

  async start(ctx: StartSessionContext, command: string, env: NodeJS.ProcessEnv): Promise<string> {
    if (ctx.resumeSessionId && !validNativeId(ctx.resumeSessionId)) throw new Error("Invalid Copilot resume session ID");
    const { CopilotClient, RuntimeConnection, ToolSet } = await import("@github/copilot-sdk");
    const web = ctx.mcp.env.EXCEL_WEB_SEARCH === "1";
    const names = exposedToolNames(ctx.mcp.env.EXCEL_ALLOW_OFFICE_JS === "1", { playwright: ctx.mcp.env.EXCEL_PLAYWRIGHT === "1" });
    const tools = new ToolSet().addBuiltIn("ask_user");
    for (const name of names) tools.addMcp(`excel-${name}`);
    if (web) tools.addBuiltIn(["web_search", "web_fetch"]);
    const cliPath = resolveCopilotPath(command, env);
    const client = new CopilotClient({
      connection: RuntimeConnection.forStdio({ path: cliPath, args: ["--no-custom-instructions", "--no-remote",
        ...(web ? [] : ["--disable-builtin-mcps"])] }),
      workingDirectory: ensureWorkspaceDir(), env, logLevel: "error",
    });
    let state: NativeSession | undefined;
    const config: SessionConfig = {
      model: ctx.model ?? undefined, reasoningEffort: ctx.effort, streaming: true, availableTools: tools,
      workingDirectory: ensureWorkspaceDir(),
      ...(ctx.systemPrompt?.trim() ? { systemMessage: { mode: "append", content: ctx.systemPrompt } } : {}),
      mcpServers: { excel: { type: "local", command: ctx.mcp.command, args: ctx.mcp.args, env: ctx.mcp.env, tools: names } },
      onPermissionRequest: request => (request.kind === "mcp" && request.serverName === "excel") || (web && request.kind === "url")
        ? { kind: "approve-once" } : { kind: "reject" },
      onUserInputRequest: async request => {
        const active = state?.active;
        if (!active || active.signal.aborted || active.controller.signal.aborted) throw new Error("Question cancelled");
        const response = await askProviderQuestions(UserInputRequestSchema.parse({ questions: [{
          id: "answer", question: request.question, options: request.choices?.map(label => ({ label })) ?? [],
          allowFreeform: request.allowFreeform ?? true,
        }] }), ctx.requestUserInput, AbortSignal.any([active.signal, active.controller.signal]));
        // The native legacy tool has no skip field. Reject its request instead of inventing an answer.
        if (response.cancelled) throw new Error("The user cancelled the question. Do not assume an answer.");
        const answer = response.answers.answer![0]!;
        return { answer, wasFreeform: !request.choices?.includes(answer) };
      },
    };
    try {
      await client.start();
      ctx.signal.throwIfAborted();
      const session = ctx.resumeSessionId ? await client.resumeSession(ctx.resumeSessionId, config) : await client.createSession(config);
      if (ctx.resumeSessionId && session.sessionId !== ctx.resumeSessionId) throw new Error("Copilot resumed a different session");
      const reader = new CopilotUsageReader(path.join(env.COPILOT_HOME ?? path.join(os.homedir(), ".copilot"),
        "session-state", session.sessionId, "events.jsonl"), session.sessionId);
      state = { client, session, reader, usage: new CumulativeUsage(), baselinePending: false };
      if (ctx.resumeSessionId) {
        const baseline = await reader.snapshot().catch(() => null);
        if (baseline) state.usage.next(baseline);
        else state.baselinePending = true;
      }
      const id = `copilot-native-${randomUUID()}`;
      this.sessions.set(id, state);
      return id;
    } catch (error) { await client.stop(); throw error; }
  }

  async *send(input: SendTurnInput, env: NodeJS.ProcessEnv): AsyncIterable<ProviderEvent> {
    const state = this.sessions.get(input.providerSessionId)!;
    const { requestId } = input;
    const queue = new AsyncQueue<ProviderEvent>();
    const controller = new AbortController();
    state.active = { signal: input.signal, controller, queue };
    const info = async (): Promise<ProviderEvent> => ({ type: "session_info", requestId, provider: "copilot",
      providerSessionId: state.session.sessionId, title: await copilotTitle(state.session.sessionId, env) });
    yield await info();
    yield { type: "turn_started", requestId };
    if (input.signal.aborted) { state.active = undefined; yield { type: "turn_finished", requestId, reason: "cancelled" }; return; }
    const streamed = new Set<string>();
    const toolNames = new Map<string, string>();
    let failed = false;
    const unsubscribe = state.session.on(event => {
      if (event.agentId || ("parentToolCallId" in event.data && event.data.parentToolCallId)) return;
      switch (event.type) {
        case "assistant.message_delta":
          streamed.add(event.data.messageId);
          queue.push({ type: "message_delta", requestId, text: event.data.deltaContent });
          break;
        case "assistant.message":
          if (!streamed.has(event.data.messageId) && event.data.content) queue.push({ type: "message_delta", requestId, text: event.data.content });
          queue.push({ type: "message_completed", requestId, text: event.data.content });
          break;
        case "tool.execution_start": {
          const tool = event.data.mcpServerName === "excel" ? event.data.mcpToolName ?? event.data.toolName : event.data.toolName;
          toolNames.set(event.data.toolCallId, tool);
          queue.push({ type: "tool_started", requestId, callId: event.data.toolCallId, tool, args: event.data.arguments ?? {} });
          break;
        }
        case "tool.execution_complete":
          queue.push({ type: "tool_finished", requestId, callId: event.data.toolCallId, tool: toolNames.get(event.data.toolCallId) ?? "unknown",
            ok: event.data.success, error: event.data.success ? null : event.data.error?.message ?? "Tool failed" });
          break;
        case "session.error":
          failed = true;
          queue.push({ type: "error", requestId, code: /auth|login/i.test(event.data.errorType) ? "login_required" : "provider_error", message: event.data.message });
          break;
      }
    });
    const cancel = () => { controller.abort(); void state.session.abort().catch(() => undefined); };
    input.signal.addEventListener("abort", cancel, { once: true });
    const context = buildTurnPrompt(input.prompt, input.selection);
    const prompt = input.userPrompt?.trim() ? `${input.userPrompt}\n\n<addin_context>\n${context}\n</addin_context>` : context;
    void state.session.sendAndWait({ prompt }, 30 * 60_000).then(async () => {
      queue.push(await info());
      const total = await state.reader.snapshot().catch(() => null);
      if (total) {
        const usage = state.usage.next(total);
        if (!state.baselinePending) queue.push({ type: "usage_update", requestId, usage });
        state.baselinePending = false;
      }
      queue.push({ type: "turn_finished", requestId, reason: input.signal.aborted ? "cancelled" : failed ? "error" : "completed" });
      queue.end();
    }).catch(async error => {
      if (!input.signal.aborted) queue.push({ type: "error", requestId, code: "provider_error", message: error instanceof Error ? error.message : String(error) });
      await state.session.abort().catch(() => undefined);
      queue.push({ type: "turn_finished", requestId, reason: input.signal.aborted ? "cancelled" : "error" });
      queue.end();
    });
    try { for await (const event of queue) yield event; }
    finally {
      controller.abort(); unsubscribe(); input.signal.removeEventListener("abort", cancel);
      if (state.active?.queue === queue) state.active = undefined;
    }
  }

  async cancel(id: string) { const state = this.sessions.get(id); state?.active?.controller.abort(); await state?.session.abort(); }
  async close(id: string) {
    const state = this.sessions.get(id);
    this.sessions.delete(id);
    if (!state) return;
    state.active?.controller.abort(); state.active?.queue.end();
    try { await state.session.disconnect(); } finally { await state.client.stop(); }
  }
  async dispose() { await Promise.all([...this.sessions.keys()].map(id => this.close(id))); }
}

/** The SDK needs an actual file path, while the add-in setting also accepts PATH commands. */
export function resolveCopilotPath(command: string, env: NodeJS.ProcessEnv): string {
  const extensions = process.platform === "win32" ? (env.PATHEXT ?? ".EXE;.CMD;.BAT").split(";") : [""];
  const candidates = path.isAbsolute(command) || command.includes(path.sep) ? [command]
    : (env.PATH ?? "").split(path.delimiter).flatMap(directory => extensions.map(extension => path.join(directory, command + extension)));
  for (const candidate of candidates) {
    if (!fs.existsSync(candidate)) continue;
    let resolved = fs.realpathSync(candidate);
    // npm's Windows command shim cannot be spawned without a shell; use its JS entrypoint.
    if (/\.(cmd|bat)$/i.test(resolved)) {
      const loader = path.join(path.dirname(resolved), "node_modules", "@github", "copilot", "npm-loader.js");
      if (fs.existsSync(loader)) resolved = loader;
    }
    if (path.basename(resolved) === "npm-loader.js") {
      const require = createRequire(resolved);
      for (const platform of process.platform === "linux" ? ["linux", "linuxmusl"] : [process.platform]) {
        try { return require.resolve(`@github/copilot-${platform}-${process.arch}`); } catch { /* try next installed platform */ }
      }
    }
    return resolved;
  }
  throw new Error(`Copilot CLI executable not found: ${command}`);
}
