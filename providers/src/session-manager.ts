import { providerSessionKey, type ProviderEvent, type ProviderId, type ToolCapabilities } from "../../protocol/src/index.js";
import type { ProviderRegistry } from "./registry.js";
import type { SendTurnInput, StartSessionContext } from "./types.js";
/** Transport-independent state shared by workbook and API conversations. */
export interface ProviderSessionState {
  id: string;
  provider: ProviderId | null;
  providerSessionId: string | null;
  providerNativeSessionId: string | null;
  providerChatSessionId: string | null;
  providerModel: string | null;
  providerEffort: "low" | "medium" | "high";
  systemPromptRevision: string | null;
  capabilities: ToolCapabilities;
}
interface SessionLogger {
  debug(message: string, details?: Record<string, unknown>): void;
  warn(message: string, details?: Record<string, unknown>): void;
}

export interface ProviderSessionSettings {
  model: string | null;
  effort: "low" | "medium" | "high";
  chatSessionId?: string;
  resumeSessionId?: string;
  capabilities: ToolCapabilities;
  systemPromptRevision: string;
}

/** Owns runtime reuse and native conversation identity for every provider.
 * Callers hold transport-independent state; adapters own transport details only.
 */
export class ProviderSessionManager<S extends ProviderSessionState = ProviderSessionState> {
  private readonly generations = new WeakMap<S, number>();
  private readonly lifetimes = new WeakMap<S, { closed: boolean }>();

  constructor(
    private readonly providers: ProviderRegistry,
    private readonly releaseTools: (session: S) => Promise<void>,
    private readonly log: SessionLogger,
  ) {}

  async ensure(
    session: S,
    settings: ProviderSessionSettings,
    context: () => StartSessionContext,
  ): Promise<string> {
    if (!session.provider) throw new Error("Select a provider before opening a session");
    const adapter = this.providers.require(session.provider);
    const sameChat = settings.resumeSessionId
      ? settings.resumeSessionId === session.providerNativeSessionId
      : !settings.chatSessionId || !session.providerChatSessionId || settings.chatSessionId === session.providerChatSessionId;
    // A configuration change replaces the runtime, while retaining the native
    // conversation. New chat is the only ordinary transition to a fresh one.
    const resumeSessionId = settings.resumeSessionId ?? (sameChat ? session.providerNativeSessionId ?? undefined : undefined);
    const changed = !sameChat
      || session.providerModel !== settings.model
      || session.providerEffort !== settings.effort
      || session.systemPromptRevision !== settings.systemPromptRevision
      || session.capabilities.webSearch !== settings.capabilities.webSearch
      || session.capabilities.playwright !== settings.capabilities.playwright;
    const runtimeClosed = this.lifetimes.get(session)?.closed === true;
    if (session.providerSessionId && (changed || runtimeClosed)) {
      await this.close(session, runtimeClosed ? "runtime_closed" : sameChat ? "configuration_change" : "chat_change", sameChat);
    }
    session.capabilities = { ...settings.capabilities };
    if (session.providerSessionId) {
      this.log.debug("runtime_reused", { sessionId: session.id, provider: adapter.id });
      return session.providerSessionId;
    }

    const generation = this.generations.get(session) ?? 0;
    const ctx = context();
    ctx.signal.throwIfAborted();
    const lifetime = { closed: false };
    const started = performance.now();
    this.log.debug("runtime_opening", { sessionId: session.id, provider: adapter.id, resume: Boolean(resumeSessionId) });
    const id = await adapter.startSession({ ...ctx, resumeSessionId, onRuntimeClosed: () => { lifetime.closed = true; } });
    // Disconnect/provider change can happen while an adapter is initializing.
    // Never attach its late result to a different conversation/provider lifecycle.
    if (lifetime.closed || ctx.signal.aborted || generation !== (this.generations.get(session) ?? 0) || session.provider !== adapter.id) {
      await adapter.closeSession(id);
      throw new Error("Provider session opening was cancelled");
    }
    this.log.debug("runtime_opened", { sessionId: session.id, provider: adapter.id, resume: Boolean(resumeSessionId), durationMs: Math.round(performance.now() - started) });
    session.providerSessionId = id;
    this.lifetimes.set(session, lifetime);
    session.providerNativeSessionId = resumeSessionId ?? null;
    session.providerChatSessionId = resumeSessionId ? providerSessionKey(adapter.id, resumeSessionId) : settings.chatSessionId ?? session.id;
    session.providerModel = settings.model;
    session.providerEffort = settings.effort;
    session.systemPromptRevision = settings.systemPromptRevision;
    return id;
  }

  async *sendTurn(session: S, input: Omit<SendTurnInput, "providerSessionId">): AsyncIterable<ProviderEvent> {
    if (!session.provider || !session.providerSessionId) throw new Error("Provider session is not open");
    const adapter = this.providers.require(session.provider);
    const id = session.providerSessionId;
    let terminal: Extract<ProviderEvent, { type: "turn_finished" }> | undefined;
    try {
      for await (const event of adapter.sendTurn({ ...input, providerSessionId: id })) {
        if (event.requestId !== input.requestId) throw new Error("Provider returned an event for a different request");
        if (event.type === "session_info" && session.providerSessionId === id && event.provider === session.provider) {
          session.providerNativeSessionId = event.providerSessionId;
          session.providerChatSessionId = providerSessionKey(event.provider, event.providerSessionId);
        }
        if (event.type === "turn_finished") terminal ??= event;
        else yield event;
      }
    } catch (error) {
      await this.dropFailedRuntime(session, id);
      throw error;
    }
    if (terminal?.reason === "error" || !terminal || this.lifetimes.get(session)?.closed) await this.dropFailedRuntime(session, id);
    // A terminal event means callers may immediately submit another turn.
    // Wait for iterator cleanup before exposing it, uniformly across adapters.
    yield terminal ?? { type: "turn_finished", requestId: input.requestId, reason: input.signal.aborted ? "cancelled" : "error" };
  }

  async cancel(session: S): Promise<void> {
    if (session.provider && session.providerSessionId) await this.providers.get(session.provider)?.cancelTurn(session.providerSessionId);
  }

  private async dropFailedRuntime(session: S, id: string): Promise<void> {
    if (session.providerSessionId !== id) return;
    await this.close(session, "turn_error", true);
  }

  async close(session: S, reason: string, preserveConversation = false): Promise<void> {
    this.generations.set(session, (this.generations.get(session) ?? 0) + 1);
    const provider = session.provider;
    const id = session.providerSessionId;
    this.lifetimes.delete(session);
    session.providerSessionId = null;
    if (!preserveConversation) {
      session.providerNativeSessionId = null;
      session.providerChatSessionId = null;
    }
    session.providerModel = null;
    session.systemPromptRevision = null;
    try { await this.releaseTools(session); }
    catch (error) { this.log.warn("provider tool cleanup failed", { error: String(error) }); }
    if (!provider || !id) return;
    try {
      await this.providers.get(provider)?.closeSession(id);
      this.log.debug("provider session closed", { sessionId: session.id, provider, reason });
    } catch (error) {
      this.log.warn("closeSession failed", { error: String(error) });
    }
  }
}
