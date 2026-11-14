import { sessionMcpEndpoint } from "./session-mcp.js";
import { createHash } from "node:crypto";
import path from "node:path";
import { parseNativeCommand, type NativeCommand } from "../../protocol/src/commands.js";
import { providerSessionKey, type AnySelectionSnapshot, type ProviderId, type ServerMessage, type ToolCapabilities } from "../../protocol/src/index.js";
import { hostSystemPrompt } from "../../protocol/src/system-prompts.js";
import type { ProviderRegistry } from "../../providers/src/registry.js";
import type { ProviderAdapter } from "../../providers/src/types.js";
import type { ProviderSessionManager } from "../../shared/src/session-manager.js";
import type { OfficeRuntime } from "./office-runtime.js";
import { attachmentPath } from "./attachments.js";
import type { BridgeConfig } from "./config.js";
import type { Delegations } from "./delegations.js";
import type { Logger } from "./logger.js";
import type { ReferenceFiles } from "./reference-files.js";
import type { SavedReferences } from "./saved-references.js";
import type { WorkbookSession } from "./session.js";
import { type SkillStore, expandSkillPrompt, projectTemplateReminder } from "./skills.js";
import type { SystemPromptStore } from "./system-prompt.js";
import type { UsageRecord, UsageStore } from "./usage.js";

export interface TurnRunnerDeps {
  isUpdating?: () => boolean;
  config: BridgeConfig;
  logger: Logger;
  providers: ProviderRegistry;
  providerSessions: ProviderSessionManager<WorkbookSession>;
  delegations: Delegations;
  externalTools: OfficeRuntime["externalTools"];
  usage: UsageStore;
  skills: SkillStore;
  systemPrompt: SystemPromptStore;
  savedReferences: SavedReferences;
  references: ReferenceFiles;
  /** Discovers the enabled external MCP tools for the pane and records them on its session. */
  discoverMcp: (session: WorkbookSession) => Promise<WorkbookSession["externalMcpTools"]>;
  /** Whether another file is open next to this pane's, which is what the tools for work across files are offered for. */
  hasOtherFiles: (session: WorkbookSession) => boolean;
}

/** What a `prompt` message from the pane carries. */
export interface PromptInput {
  requestId: string;
  prompt: string;
  selection: AnySelectionSnapshot | null;
  context?: string;
  attachmentIds?: string[];
  model?: string | null;
  referencePaths?: string[];
  effort?: "low" | "medium" | "high";
  chatSessionId?: string;
  resumeSessionId?: string;
  capabilities?: ToolCapabilities;
  /** A provider slash command rather than a chat message. */
  isCommand?: boolean;
  /** Set when another file's assistant asked for this turn. */
  delegationId?: string;
}

/** What a provider runtime is opened with: the pane's choices for its next prompt. */
export interface RuntimeInput {
  model: string | null;
  effort: "low" | "medium" | "high";
  chatSessionId?: string;
  resumeSessionId?: string;
  capabilities: ToolCapabilities;
}

/**
 * One turn of the conversation: build the prompt, make sure the provider session exists, stream its events to the pane,
 * record usage, and finish exactly once. Also the provider-level actions around a turn: switching, cancelling, closing.
 *
 * The provider's runtime lives as long as the pane's connection: it is opened ahead of the first message (`warm`) and
 * reused by every prompt after it, so a message only pays for the model's answer.
 */
export class TurnRunner {
  private readonly activeUsage = new Map<string, UsageRecord>();
  /** The runtime each pane is being opened for ahead of a prompt, with what it is opened for. */
  private readonly warming = new Map<string, { controller: AbortController; key: string }>();
  /** What each pane last said its next prompt runs with, to open the runtime again when something other than the pane changes the tools. */
  private readonly lastInputs = new WeakMap<WorkbookSession, RuntimeInput>();

  constructor(private readonly deps: TurnRunnerDeps) {}

  /**
   * Another file opened or closed: the tools for work across files are offered only while there is one. If the runtime is open
   * and idle, it is opened again with the tools that now apply, now rather than when the next message is sent.
   */
  refreshTools(session: WorkbookSession): void {
    const input = this.lastInputs.get(session);
    if (!input || !session.providerSessionId || session.isBusy) return;
    if (this.deps.hasOtherFiles(session) === Boolean(session.capabilities.connectedFiles)) return;
    void this.warm(session, input);
  }

  /** Open the provider's runtime now so the next prompt finds it ready. Never throws: a prompt that needs the runtime reports its own failure. */
  async warm(session: WorkbookSession, input: RuntimeInput): Promise<void> {
    if (!session.provider || session.isBusy || this.deps.isUpdating?.()) return;
    const adapter = this.deps.providers.get(session.provider);
    if (!adapter) return;
    this.lastInputs.set(session, input);
    const key = JSON.stringify([session.provider, input]);
    const current = this.warming.get(session.id);
    if (current?.key === key) return;
    current?.controller.abort();
    const entry = { controller: new AbortController(), key };
    this.warming.set(session.id, entry);
    try {
      const detection = await adapter.detect();
      if (entry.controller.signal.aborted || session.isBusy || !detection.installed || detection.status !== "ready") return;
      await this.openRuntime(session, adapter, input, this.promptSettings(session), entry.controller.signal);
    } catch (error) {
      if (!entry.controller.signal.aborted) this.deps.logger.debug("runtime warm-up failed", { error: String(error) });
    } finally {
      if (this.warming.get(session.id) === entry) this.warming.delete(session.id);
    }
  }

  private promptSettings(session: WorkbookSession) {
    const stored = this.deps.systemPrompt.forHost(session.host).read();
    // Excel keeps the user's prompt as it is; Word and PowerPoint add their own rules (and a new revision, so a running session restarts with them).
    if (session.host === "excel") return stored;
    const content = hostSystemPrompt(session.host, stored.content, { connectedFiles: this.deps.hasOtherFiles(session) });
    return { ...stored, content, revision: createHash("sha256").update(`${stored.revision}:${session.host}:${content}`).digest("hex") };
  }

  /** Make sure the pane's provider runtime is open with these settings; an open one that already has them is reused. */
  private async openRuntime(
    session: WorkbookSession, adapter: ProviderAdapter, input: RuntimeInput, promptSettings: { content: string; revision: string },
    signal: AbortSignal, onOpening?: () => void,
  ): Promise<void> {
    const externalCatalog = await this.deps.discoverMcp(session);
    const catalogRevision = createHash("sha256").update(JSON.stringify([...externalCatalog.values()].map(tool => tool.definition))).digest("hex");
    const { model, effort, chatSessionId, resumeSessionId } = input;
    await this.deps.providerSessions.ensure(session, {
      model, effort, chatSessionId, resumeSessionId,
      capabilities: { ...input.capabilities, webSearch: input.capabilities.webSearch && adapter.capabilities.nativeWebSearch,
        connectedFiles: this.deps.hasOtherFiles(session) },
      systemPromptRevision: `${promptSettings.revision}:${catalogRevision}`,
    }, () => ({
      requestUserInput: (request, requestSignal) => session.requestUserInput(request, requestSignal),
      authorizeTool: async (args) => {
        const result = await session.requestTool("__authorize_provider_tool", args, 120_000);
        return result.ok && (result.data as { approved?: boolean })?.approved === true;
      },
      workbookConnectionId: session.workbookConnectionId,
      sessionId: session.id,
      model,
      effort,
      resumeSessionId,
      mcp: sessionMcpEndpoint(this.deps.config, session, adapter.capabilities),
      signal,
      systemPrompt: promptSettings.content,
    }), onOpening);
  }

  async setProvider(session: WorkbookSession, provider: ProviderId): Promise<void> {
    const adapter = this.deps.providers.get(provider);
    if (!adapter) {
      return void session.send({
        type: "error",
        requestId: null,
        code: "unknown_provider",
        message: `provider '${provider}' is not available`,
      });
    }

    if (session.provider && session.provider !== provider) {
      await this.closeProviderProcess(session, "provider_change");
    }
    session.provider = provider;

    const detection = await adapter.detect(true);
    session.send({ type: "provider_status", provider, detection });
  }

  async run(session: WorkbookSession, input: PromptInput): Promise<void> {
    if (this.deps.isUpdating?.()) {
      session.send({ type: "error", requestId: input.requestId, code: "maintenance", message: "Office Partner is switching releases. Try again after it reconnects." });
      session.send({ type: "turn_finished", requestId: input.requestId, reason: "error" });
      return;
    }
    const { requestId, attachmentIds, referencePaths, model, chatSessionId, resumeSessionId, delegationId, context,
      effort = "medium", isCommand = false } = input;
    let { prompt, selection, capabilities = { webSearch: false, playwright: false } } = input;
    const userPrompt = prompt;
    const { delegations, references, skills, logger, providerSessions } = this.deps;
    if (!session.provider) {
      if (delegationId) delegations.decline(session, delegationId, "That file has no AI provider selected yet.");
      return void session.send({
        type: "error",
        requestId,
        code: "no_provider",
        message: "select a provider before sending a prompt",
      });
    }
    if (session.isBusy) {
      if (delegationId) delegations.decline(session, delegationId, "That file is busy with another request.");
      return void session.send({
        type: "error",
        requestId,
        code: "busy",
        message: "another turn is already running for this workbook",
      });
    }

    if (attachmentIds !== undefined && (!Array.isArray(attachmentIds) || attachmentIds.length > 10 || attachmentIds.some((id) => typeof id !== "string" || !/^[0-9a-f-]{36}$/.test(id)))) {
      session.send({ type: "error", requestId, code: "invalid_attachment", message: "Invalid attached file reference." });
      session.send({ type: "turn_finished", requestId, reason: "error" });
      return;
    }
    if (referencePaths?.length) {
      try {
        prompt += references.mentionedFilesBlock(referencePaths, session.referenceGrants);
      } catch (error) {
        session.send({ type: "error", requestId, code: "invalid_attachment", message: error instanceof Error ? error.message : String(error) });
        session.send({ type: "turn_finished", requestId, reason: "error" });
        return;
      }
    }
    if (attachmentIds?.length) {
      const attached = attachmentIds.map((id) => session.attachments.get(id));
      if (attached.some((file) => !file)) {
        session.send({ type: "error", requestId, code: "attachment_missing", message: "Attached file is unavailable; select it again." });
        session.send({ type: "turn_finished", requestId, reason: "error" });
        return;
      }
      prompt += "\n\n<user_attached_files>\n" + attached.map((file) =>
        JSON.stringify({ name: file!.name, path: attachmentPath(file!.id), kind: file!.kind }),
      ).join("\n") + "\n</user_attached_files>";
    }

    const adapter = this.deps.providers.require(session.provider);
    const selectedModel = model ?? null;
    const promptSettings = this.promptSettings(session);
    if (delegationId && !delegations.accept(session, delegationId, requestId)) {
      // Nobody waits for this any more (it timed out, the asker stopped, or it was never asked): do not change the file for no one.
      session.send({ type: "error", requestId, code: "cancelled", message: "The other file's request is no longer waiting for an answer, so it was not run." });
      session.send({ type: "turn_finished", requestId, reason: "cancelled" });
      return;
    }
    const controller = session.beginTurn(requestId);
    this.activeUsage.set(`${session.id}:${requestId}`, {
      sessionId: resumeSessionId ? providerSessionKey(session.provider, resumeSessionId) : chatSessionId ?? session.id,
      providerSessionId: resumeSessionId, requestId, provider: session.provider,
      model: selectedModel, startedAt: Date.now(), durationMs: 0, reason: "completed", usage: null,
    });
    // `/skill-name request` loads the skill's instructions into this turn.
    const expanded = isCommand ? { prompt, skill: undefined } : expandSkillPrompt(prompt, skills);
    if (expanded.skill) logger.info("skill invoked", { skill: expanded.skill });
    prompt = isCommand ? userPrompt : [resumeSessionId ? undefined : context, projectTemplateReminder(skills), this.deps.savedReferences.prompt(), expanded.prompt].filter(Boolean).join("\n\n");

    // Where a slow message spends its time: before the provider is asked (opening the runtime) or inside it.
    const timing = { start: performance.now(), ready: 0, firstEvent: 0, firstText: 0, tools: 0 };
    const since = () => Math.round(performance.now() - timing.start);
    try {
      controller.signal.throwIfAborted();
      // An open runtime needs no check: only a prompt that has to open one asks whether the CLI is there.
      if (!session.providerSessionId) {
        const detection = await adapter.detect();
        if (!detection.installed && !resumeSessionId && !session.providerNativeSessionId) throw new Error(detection.hint ?? "Provider CLI is not installed.");
      }
      let command: NativeCommand | undefined;
      if (isCommand) {
        const parsed = parseNativeCommand(userPrompt);
        const catalog = await adapter.listCommands?.(session.providerSessionId ?? undefined) ?? [];
        const entry = catalog.find(item => item.name === parsed?.name);
        if (!parsed || entry?.execution !== "native") throw new Error(entry?.reason ?? "This command is not available through the provider. Refresh the / menu.");
        if (attachmentIds?.length || referencePaths?.length) throw new Error("Send attached files in a regular message, separately from a command.");
        command = parsed;
      }
      controller.signal.throwIfAborted();
      // A warm-up still in progress is what this prompt waits for: say so, as when the runtime is opened for the prompt itself.
      if (this.warming.has(session.id)) session.send({ type: "session_starting", requestId });
      const runtimeInput: RuntimeInput = { model: selectedModel, effort, chatSessionId, resumeSessionId, capabilities };
      this.lastInputs.set(session, runtimeInput);
      await this.openRuntime(session, adapter, runtimeInput, promptSettings, controller.signal, () => session.send({ type: "session_starting", requestId }));

      timing.ready = since();
      for await (const event of providerSessions.sendTurn(session, {
        systemPrompt: promptSettings.content,
        requestId,
        model: selectedModel,
        effort,
        prompt,
        command,
        userPrompt,
        selection,
        signal: controller.signal,
      })) {
        timing.firstEvent ||= since();
        if (event.type === "message_delta" || event.type === "message_completed") timing.firstText ||= since();
        if (event.type === "tool_started") timing.tools++;
        if (event.type === "turn_finished") {
          this.finishTurn(session, requestId, event.reason);
        } else if (!session.hasFinished(requestId)) {
          delegations.observe(session, requestId, event);
          if (event.type === "session_info" && event.requestId === requestId && event.provider === session.provider) {
            const record = this.activeUsage.get(`${session.id}:${requestId}`);
            if (record) {
              record.sessionId = providerSessionKey(event.provider, event.providerSessionId);
              record.providerSessionId = event.providerSessionId;
              if (event.title) record.title = event.title;
            }
          }
          if (event.type === "usage_update") {
            const record = this.activeUsage.get(`${session.id}:${requestId}`);
            if (record) record.usage = event.usage;
          }
          session.send(event as ServerMessage);
        }
      }
      logger.info("turn timing", { provider: session.provider, ...timing, start: undefined, total: since() });
      this.finishTurn(session, requestId, controller.signal.aborted ? "cancelled" : "completed");
    } catch (error) {
      const aborted = controller.signal.aborted;
      if (!session.hasFinished(requestId)) {
        const failure = {
          type: "error" as const,
          requestId,
          code: aborted ? "cancelled" : resumeSessionId && !session.providerSessionId ? "resume_failed" : "provider_error",
          message: aborted ? "turn cancelled" : `${resumeSessionId && !session.providerSessionId ? "Could not resume the saved CLI session. No replacement session was created. " : ""}${String(error instanceof Error ? error.message : error)}`,
        };
        delegations.observe(session, requestId, failure);
        session.send(failure);
      }
      this.finishTurn(session, requestId, aborted ? "cancelled" : "error");
    } finally {
      session.endTurn(requestId);
      session.touch();
    }
  }

  /** Emit `turn_finished` exactly once per request. */
  private finishTurn(
    session: WorkbookSession,
    requestId: string,
    reason: "completed" | "cancelled" | "error",
  ): void {
    if (!session.markFinished(requestId)) return;
    this.deps.delegations.finish(session, requestId, reason);
    const key = `${session.id}:${requestId}`;
    const record = this.activeUsage.get(key);
    if (record) {
      this.activeUsage.delete(key);
      record.durationMs = Math.max(0, Date.now() - record.startedAt);
      record.reason = reason;
      try { this.deps.usage.append(record); }
      catch (error) {
        this.deps.logger.error("could not save usage", { error: String(error) });
        session.send({ type: "error", requestId, code: "usage_save_failed", message: "Usage statistics could not be saved to disk." });
      }
    }
    // Release the workbook before telling the pane, otherwise a client that
    // sends its next prompt on `turn_finished` races the adapter's cleanup
    // and gets a spurious `busy`.
    session.takeFollowups(requestId, "next_turn");
    session.endTurn(requestId);
    session.send({ type: "turn_finished", requestId, reason });
  }

  /** Give a message to the turn that is running, through the provider; false when it must be queued instead. */
  steer(session: WorkbookSession, requestId: string, text: string): Promise<boolean> {
    if (session.activeRequestId !== requestId || session.turnController?.signal.aborted) return Promise.resolve(false);
    return this.deps.providerSessions.steer(session, text);
  }

  async cancel(session: WorkbookSession, requestId: string): Promise<void> {
    if (session.activeRequestId !== requestId) return;
    session.abortTurn();
    void this.deps.externalTools.closeSession(session.id, "browser");
    try {
      await this.deps.providerSessions.cancel(session);
    } catch (error) {
      this.deps.logger.warn("cancelTurn failed", { error: String(error) });
    }
    // The adapter normally ends its stream right after cancel. If it hangs,
    // report the cancellation anyway and drop the provider process so the
    // next prompt starts clean instead of hitting `busy` forever.
    const watchdog = setTimeout(() => {
      if (session.activeRequestId !== requestId) return;
      this.deps.logger.warn("provider ignored cancel; closing its process", { requestId });
      this.finishTurn(session, requestId, "cancelled");
      session.endTurn(requestId);
      void this.closeProviderProcess(session, "cancel_timeout");
    }, this.deps.config.cancelGraceMs);
    watchdog.unref();
  }

  async closeProviderProcess(session: WorkbookSession, reason: string): Promise<void> {
    if (session.activeRequestId) session.takeFollowups(session.activeRequestId, "next_turn");
    session.abortTurn();
    session.activeRequestId = null;
    this.warming.get(session.id)?.controller.abort();
    this.warming.delete(session.id);
    await this.deps.providerSessions.close(session, reason);
  }
}
