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
import { normalizeCommands } from "./commands.js";

interface NativeSession {
  ask: StartSessionContext["requestUserInput"];
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

  async listModels(command: string, env: NodeJS.ProcessEnv) {
    const { CopilotClient, RuntimeConnection } = await import("@github/copilot-sdk");
    const client = new CopilotClient({ connection: RuntimeConnection.forStdio({ path: resolveCopilotPath(command, env),
      args: ["--no-custom-instructions", "--disable-builtin-mcps"] }), workingDirectory: ensureWorkspaceDir(), env, logLevel: "error" });
    try {
      await client.start();
      return { models: (await client.listModels()).map(model => ({ id: model.id, name: model.name })), defaultModelId: null };
    } finally { await client.stop(); }
  }

  async listCommands(command: string, env: NodeJS.ProcessEnv, id?: string) {
    const current = id ? this.sessions.get(id) : undefined;
    if (current) return normalizeCommands((await current.session.rpc.commands.list()).commands);
    const { CopilotClient, RuntimeConnection } = await import("@github/copilot-sdk");
    const client = new CopilotClient({ connection: RuntimeConnection.forStdio({ path: resolveCopilotPath(command, env),
      args: ["--no-custom-instructions", "--disable-builtin-mcps"] }),
      workingDirectory: ensureWorkspaceDir(), env, logLevel: "error" });
    try {
      await client.start();
      const session = await client.createSession({ availableTools: [], onPermissionRequest: () => ({ kind: "reject" }) });
      try { return normalizeCommands((await session.rpc.commands.list()).commands); }
      finally { await session.disconnect(); }
    } finally { await client.stop(); }
  }

  async start(ctx: StartSessionContext, command: string, env: NodeJS.ProcessEnv): Promise<string> {
    if (ctx.resumeSessionId && !validNativeId(ctx.resumeSessionId)) throw new Error("Invalid Copilot resume session ID");
    const { CopilotClient, RuntimeConnection, ToolSet } = await import("@github/copilot-sdk");
    const web = ctx.mcp.env.EXCEL_WEB_SEARCH === "1";
    const names = ctx.mcp.toolNames ?? exposedToolNames(ctx.mcp.env.EXCEL_ALLOW_OFFICE_JS === "1", { playwright: ctx.mcp.env.EXCEL_PLAYWRIGHT === "1", qa: ctx.mcp.env.EXCEL_QA === "1" });
    const tools = new ToolSet();
    if (!ctx.mcp.toolNames) tools.addBuiltIn("ask_user");
    for (const name of names) tools.addMcp(`excel-${name}`);
    if (web) tools.addBuiltIn(["web_search", "web_fetch"]);
    const cliPath = resolveCopilotPath(command, env);
    const client = new CopilotClient({
      connection: RuntimeConnection.forStdio({ path: cliPath, args: ["--no-custom-instructions",
        ...(web ? [] : ["--disable-builtin-mcps"])] }),
      workingDirectory: ensureWorkspaceDir(), env, logLevel: "error",
    });
    let state: NativeSession | undefined;
    const config: SessionConfig = {
      model: ctx.model ?? undefined, streaming: true, availableTools: tools,
      // Auto selects its own reasoning settings. The CLI accepts effort at
      // creation but rejects it on resume, for both workbook and API sessions.
      ...(!ctx.model || ctx.model === "auto" ? {} : { reasoningEffort: ctx.effort }),
      workingDirectory: ensureWorkspaceDir(),
      ...(ctx.systemPrompt?.trim() ? { systemMessage: { mode: ctx.systemPromptMode ?? "append", content: ctx.systemPrompt } } : {}),
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
      state = { client, session, reader, usage: new CumulativeUsage(), baselinePending: false, ask: ctx.requestUserInput };
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
    // Streamed text not yet matched to a final `assistant.message`. Some models (for example the
    // "commentary" phase of gpt-6-luna) stream under one message id and put the final text in a later
    // message with another id, after an empty one. Matching by text keeps that from being shown twice.
    let pending = "";
    const toolNames = new Map<string, string>();
    let failed = false;
    const unsubscribe = state.session.on(event => {
      if (event.agentId || ("parentToolCallId" in event.data && event.data.parentToolCallId)) return;
      switch (event.type) {
        case "assistant.message_delta":
          streamed.add(event.data.messageId);
          pending += event.data.deltaContent;
          queue.push({ type: "message_delta", requestId, text: event.data.deltaContent });
          break;
        case "assistant.message": {
          const content = event.data.content ?? "";
          // Reasoning-only or tool-only messages carry no text; keep any pending deltas for the message that does.
          if (!content) break;
          let missing = content;
          if (streamed.has(event.data.messageId)) missing = "";
          else if (pending && content.startsWith(pending)) missing = content.slice(pending.length);
          else if (pending && pending.endsWith(content)) missing = "";
          if (missing) queue.push({ type: "message_delta", requestId, text: missing });
          streamed.add(event.data.messageId);
          pending = "";
          queue.push({ type: "message_completed", requestId, text: content });
          break;
        }
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
    const run = async () => {
      if (!input.command) { await state.session.sendAndWait({ prompt }, 30 * 60_000); return; }
      const { name } = input.command;
      let args = input.command.args;
      for (let depth = 0; depth < 5; depth++) {
        if (input.signal.aborted || controller.signal.aborted) return;
        const result = await state.session.rpc.commands.invoke({ name, input: args });
        switch (result.kind) {
          case "text": queue.push({ type: "message_completed", requestId, text: result.text }); return;
          case "completed":
            if (result.mode) await state.session.rpc.mode.set({ mode: result.mode });
            queue.push({ type: "message_completed", requestId, text: result.message ?? `/${name} completed.` }); return;
          case "add-timeline-entry": queue.push({ type: "message_completed", requestId, text: result.entry.text }); return;
          case "agent-prompt":
            if (result.notice) queue.push({ type: "message_completed", requestId, text: result.notice });
            await state.session.sendAndWait({ prompt: result.prompt, agentMode: result.mode }, 30 * 60_000); return;
          case "select-subcommand": {
            const answer = await askProviderQuestions(UserInputRequestSchema.parse({ questions: [{ id: "subcommand", question: result.title,
              options: result.options.slice(0, 30).map(option => ({ label: option.name, description: option.description })), allowFreeform: true }] }),
            state.ask, AbortSignal.any([input.signal, controller.signal]));
            if (answer.cancelled) { queue.push({ type: "message_completed", requestId, text: "Command cancelled." }); return; }
            args = `${args}${args ? " " : ""}${answer.answers.subcommand![0]}`;
            break;
          }
          default: throw new Error(`/${name} requires a CLI dialog. Use the add-in's corresponding setting instead.`);
        }
      }
      throw new Error(`/${name} needs more arguments. Choose a subcommand explicitly.`);
    };
    void run().then(async () => {
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
  const suffixes = (env.PATHEXT ?? ".EXE;.CMD;.BAT").split(";");
  // npm also installs a Unix shell script without an extension. Windows must prefer its .cmd shim.
  const extensions = process.platform === "win32" ? path.extname(command) ? ["", ...suffixes] : [...suffixes, ""] : [""];
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
