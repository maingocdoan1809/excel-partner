import { sessionMcpEndpoint } from "./session-mcp.js";
import { createHash } from "node:crypto";
import path from "node:path";
import { parseNativeCommand, type NativeCommand } from "../../protocol/src/commands.js";
import { providerSessionKey, supportsNativeWebSearch, type AnySelectionSnapshot, type ProviderId, type ServerMessage, type ToolCapabilities } from "../../protocol/src/index.js";
import { hostSystemPrompt } from "../../providers/src/host-prompts.js";
import type { ProviderRegistry } from "../../providers/src/registry.js";
import type { ProviderSessionManager } from "../../providers/src/session-manager.js";
import type { OfficeRuntime } from "../../runtime/src/office-runtime.js";
import { attachmentPath } from "./attachments.js";
import type { BridgeConfig } from "./config.js";
import type { Delegations } from "./delegations.js";
import type { Logger } from "./logger.js";
import type { ReferenceFiles } from "./reference-files.js";
import type { SavedReferences } from "./saved-references.js";
import { enrichSelectionForHost } from "./selection.js";
import type { WorkbookSession } from "./session.js";
import { type SkillStore, expandSkillPrompt, projectTemplateReminder } from "./skills.js";
import type { SystemPromptStore } from "./system-prompt.js";
import type { UsageRecord, UsageStore } from "./usage.js";
import type { SelectionEnricher } from "./windows-selection.js";

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
  enrichSelection: SelectionEnricher;
  /** Discovers the enabled external MCP tools for the pane and records them on its session. */
  discoverMcp: (session: WorkbookSession) => Promise<WorkbookSession["externalMcpTools"]>;
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

/**
 * One turn of the conversation: build the prompt, make sure the provider session exists, stream its events to the pane,
 * record usage, and finish exactly once. Also the provider-level actions around a turn: switching, cancelling, closing.
 */
export class TurnRunner {
  private readonly activeUsage = new Map<string, UsageRecord>();

  constructor(private readonly deps: TurnRunnerDeps) {}

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
        const selected = referencePaths.map((file) => references.mentionedPath(file, session.referenceGrants));
        prompt += "\n\n<user_attached_files>\n" + selected.map(({ path: file, folder }) => JSON.stringify(folder
          ? { name: path.basename(file), path: file, kind: "reference_folder", hint: "A folder the user pointed at. Call list_reference_files with this path to see what is inside, then read the files you need." }
          : { name: path.basename(file), path: file, kind: "reference" }),
        ).join("\n") + "\n</user_attached_files>";
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

    capabilities = { ...capabilities, webSearch: capabilities.webSearch && supportsNativeWebSearch(session.provider) };
    const selectedModel = model ?? null;
    const stored = this.deps.systemPrompt.read();
    // Excel keeps the user's prompt as it is; Word and PowerPoint add their own rules (and a new revision, so a running session restarts with them).
    const promptSettings = session.host === "excel" ? stored
      : { ...stored, content: hostSystemPrompt(session.host, stored.content), revision: createHash("sha256").update(`${stored.revision}:${session.host}:${hostSystemPrompt(session.host, stored.content)}`).digest("hex") };
    const adapter = this.deps.providers.require(session.provider);
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

    try {
      if (!isCommand && selection) selection = await enrichSelectionForHost(session.host, selection, this.deps.enrichSelection, controller.signal);
      controller.signal.throwIfAborted();
      const detection = await adapter.detect();
      if (!detection.installed) throw new Error(detection.hint ?? "Provider CLI is not installed.");
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
      const externalCatalog = await this.deps.discoverMcp(session);
      const catalogRevision = createHash("sha256").update(JSON.stringify([...externalCatalog.values()].map(tool => tool.definition))).digest("hex");
      await providerSessions.ensure(session, {
        model: selectedModel, effort, chatSessionId, resumeSessionId, capabilities,
        systemPromptRevision: `${promptSettings.revision}:${catalogRevision}`,
      }, () => ({
        requestUserInput: (request, signal) => session.requestUserInput(request, signal),
        authorizeTool: async (args) => {
          const result = await session.requestTool("__authorize_provider_tool", args, 120_000);
          return result.ok && (result.data as { approved?: boolean })?.approved === true;
        },
        workbookConnectionId: session.workbookConnectionId,
        sessionId: session.id,
        model: selectedModel,
        effort,
        resumeSessionId,
        mcp: sessionMcpEndpoint(this.deps.config, session),
        signal: controller.signal,
        systemPrompt: promptSettings.content,
      }), () => session.send({ type: "session_starting", requestId }));

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
    await this.deps.providerSessions.close(session, reason);
  }
}
