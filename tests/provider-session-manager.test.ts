import { describe, expect, it, vi } from "vitest";
import { ProviderSessionManager, type ProviderSessionSettings } from "../shared/src/session-manager.js";
import { WorkbookSession } from "../bridge/src/session.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { MockAdapter } from "../providers/src/mock.js";
import type { SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderEvent, ProviderId } from "../protocol/src/index.js";

class RuntimeAdapter extends MockAdapter {
  starts: StartSessionContext[] = [];
  closes: string[] = [];
  cancelled: string[] = [];
  cleaned = false;
  fail = false;
  override async startSession(ctx: StartSessionContext): Promise<string> {
    this.starts.push(ctx);
    if (this.fail) throw new Error("Session missing");
    return `runtime-${this.starts.length}`;
  }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    try {
      yield { type: "session_info", provider: this.id, providerSessionId: "native", requestId: input.requestId };
      yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
    } finally { this.cleaned = true; }
  }
  override async closeSession(id: string) { this.closes.push(id); }
  override async cancelTurn(id: string) { this.cancelled.push(id); }
}

function setup(provider: ProviderId) {
  const adapter = new RuntimeAdapter({ id: provider });
  const release = vi.fn(async () => {});
  const manager = new ProviderSessionManager(new ProviderRegistry().register(adapter), release, { debug() {}, warn() {} });
  const session = new WorkbookSession("workbook");
  session.provider = provider;
  const signal = new AbortController().signal;
  const settings: ProviderSessionSettings = { model: null, effort: "medium", capabilities: { webSearch: false, playwright: false },
    chatSessionId: "chat", systemPromptRevision: "revision" };
  const context = (): StartSessionContext => ({ workbookConnectionId: session.workbookConnectionId, sessionId: session.id, signal,
    mcp: { command: "node", args: [], bridgeUrl: "", sessionId: session.id, token: "", env: {} } });
  const input = { requestId: "turn", prompt: "hello", selection: null, signal };
  async function complete() { for await (const _ of manager.sendTurn(session, input)) { /* drain */ } }
  return { adapter, manager, session, settings, context, input, complete, release };
}

describe.each(["claude", "copilot", "codex", "opencode"] as const)("%s shared session lifecycle", provider => {
  it("opens once and reuses the runtime for subsequent turns", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context);
    await h.complete();
    await h.manager.ensure(h.session, { ...h.settings, chatSessionId: `${provider}:native`, resumeSessionId: "native" }, h.context);
    await h.complete();
    expect(h.adapter.starts).toHaveLength(1);
    expect(h.adapter.closes).toEqual([]);
  });

  it("opens one runtime when two callers ask at the same time", async () => {
    const h = setup(provider);
    const [first, second] = await Promise.all([
      h.manager.ensure(h.session, h.settings, h.context),
      h.manager.ensure(h.session, h.settings, h.context),
    ]);
    expect(first).toBe(second);
    expect(h.adapter.starts).toHaveLength(1);
  });

  it("announces only a runtime that is really being opened, before it opens", async () => {
    const h = setup(provider);
    const opening = vi.fn(() => { expect(h.adapter.starts).toHaveLength(0); });
    await h.manager.ensure(h.session, h.settings, h.context, opening);
    await h.complete();
    await h.manager.ensure(h.session, { ...h.settings, chatSessionId: `${provider}:native`, resumeSessionId: "native" }, h.context, opening);
    expect(opening).toHaveBeenCalledTimes(1);
    expect(h.adapter.starts).toHaveLength(1);
  });

  it.each([
    { model: "new-model" }, { effort: "high" as const }, { systemPromptRevision: "new-revision" },
    { capabilities: { webSearch: true, playwright: false } }, { capabilities: { webSearch: false, playwright: true } },
  ])("restarts for configuration changes and resumes the native conversation (%j)", async patch => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context); await h.complete();
    await h.manager.ensure(h.session, { ...h.settings, chatSessionId: `${provider}:native`, ...patch }, h.context);
    expect(h.adapter.closes).toEqual(["runtime-1"]);
    expect(h.adapter.starts[1]).toMatchObject({ resumeSessionId: "native" });
    expect(h.session.providerChatSessionId).toBe(`${provider}:native`);
    expect(h.release).toHaveBeenCalledTimes(1);
  });

  it("starts fresh for New chat and explicitly resumes a different saved chat", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context); await h.complete();
    await h.manager.ensure(h.session, { ...h.settings, chatSessionId: "new-chat" }, h.context);
    expect(h.adapter.starts[1]?.resumeSessionId).toBeUndefined();
    await h.manager.ensure(h.session, { ...h.settings, resumeSessionId: "saved" }, h.context);
    expect(h.adapter.starts[2]?.resumeSessionId).toBe("saved");
  });

  it("reports failed resume without silently opening a fresh conversation", async () => {
    const h = setup(provider); h.adapter.fail = true;
    await expect(h.manager.ensure(h.session, { ...h.settings, resumeSessionId: "missing" }, h.context)).rejects.toThrow("Session missing");
    expect(h.adapter.starts).toHaveLength(1);
    expect(h.adapter.starts[0]?.resumeSessionId).toBe("missing");
    expect(h.session.providerSessionId).toBeNull();
  });

  it("exposes completion only after adapter cleanup", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context);
    for await (const event of h.manager.sendTurn(h.session, h.input)) {
      if (event.type === "turn_finished") expect(h.adapter.cleaned).toBe(true);
    }
  });

  it("delegates cancel and clears runtime identity on close", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context); await h.complete();
    await h.manager.cancel(h.session);
    expect(h.adapter.cancelled).toEqual(["runtime-1"]);
    await h.manager.close(h.session, "disconnect");
    await h.manager.close(h.session, "disconnect");
    expect(h.adapter.closes).toEqual(["runtime-1"]);
    expect(h.session.providerSessionId).toBeNull();
    expect(h.session.providerNativeSessionId).toBeNull();
  });

  it("closes the adapter even when external tool cleanup fails", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context);
    h.release.mockRejectedValueOnce(new Error("Browser cleanup failed"));
    await h.manager.close(h.session, "disconnect");
    expect(h.adapter.closes).toEqual(["runtime-1"]);
    expect(h.session.providerSessionId).toBeNull();
  });

  it("disposes a runtime that finishes opening after disconnect", async () => {
    const h = setup(provider);
    let resolve!: (id: string) => void;
    vi.spyOn(h.adapter, "startSession").mockImplementation(() => new Promise(done => { resolve = done; }));
    const opening = h.manager.ensure(h.session, h.settings, h.context);
    await h.manager.close(h.session, "disconnect");
    resolve("late-runtime");
    await expect(opening).rejects.toThrow("cancelled");
    expect(h.adapter.closes).toEqual(["late-runtime"]);
    expect(h.session.providerSessionId).toBeNull();
  });

  it("drops a crashed runtime, retaining native identity for the next explicit turn", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context); await h.complete();
    vi.spyOn(h.adapter, "sendTurn").mockImplementation(async function* () { throw new Error("Runtime crashed"); });
    await expect(h.complete()).rejects.toThrow("Runtime crashed");
    expect(h.adapter.closes).toEqual(["runtime-1"]);
    await h.manager.ensure(h.session, { ...h.settings, chatSessionId: `${provider}:native` }, h.context);
    expect(h.adapter.starts[1]?.resumeSessionId).toBe("native");
  });

  it("resumes before sending when the runtime exits while idle", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context); await h.complete();
    h.adapter.starts[0]!.onRuntimeClosed!();
    await h.manager.ensure(h.session, { ...h.settings, resumeSessionId: "native" }, h.context);
    expect(h.adapter.closes).toEqual(["runtime-1"]);
    expect(h.adapter.starts[1]?.resumeSessionId).toBe("native");
    // A late exit callback from the old runtime cannot invalidate its replacement.
    h.adapter.starts[0]!.onRuntimeClosed!();
    await h.manager.ensure(h.session, { ...h.settings, resumeSessionId: "native" }, h.context);
    expect(h.adapter.starts).toHaveLength(2);
  });

  it("does not reuse a runtime that died during a cancelled turn", async () => {
    const h = setup(provider);
    await h.manager.ensure(h.session, h.settings, h.context); await h.complete();
    vi.spyOn(h.adapter, "sendTurn").mockImplementation(async function* (input) {
      h.adapter.starts[0]!.onRuntimeClosed!();
      yield { type: "turn_finished", requestId: input.requestId, reason: "cancelled" };
    });
    await h.complete();
    expect(h.session.providerSessionId).toBeNull();
    expect(h.session.providerNativeSessionId).toBe("native");
    await h.manager.ensure(h.session, { ...h.settings, resumeSessionId: "native" }, h.context);
    expect(h.adapter.starts[1]?.resumeSessionId).toBe("native");
  });
});

describe("a model or effort the CLI can change in the middle of a session", () => {
  class LiveAdapter extends RuntimeAdapter {
    override readonly capabilities = { ...MockAdapter.capabilities, liveModel: true, liveEffort: true };
  }
  class EffortOnlyAdapter extends RuntimeAdapter {
    override readonly capabilities = { ...MockAdapter.capabilities, liveModel: false, liveEffort: true };
  }
  function live(kind: typeof RuntimeAdapter) {
    const adapter = new kind({ id: "copilot" });
    const manager = new ProviderSessionManager(new ProviderRegistry().register(adapter), vi.fn(async () => {}), { debug() {}, warn() {} });
    const session = new WorkbookSession("workbook");
    session.provider = "copilot";
    const signal = new AbortController().signal;
    const settings: ProviderSessionSettings = { model: "a", effort: "medium", capabilities: { webSearch: false, playwright: false }, chatSessionId: "chat", systemPromptRevision: "revision" };
    const context = (): StartSessionContext => ({ workbookConnectionId: session.workbookConnectionId, sessionId: session.id, signal,
      mcp: { command: "node", args: [], bridgeUrl: "", sessionId: session.id, token: "", env: {} } });
    return { adapter, manager, session, settings, context };
  }

  it("is handed over with the next turn: the CLI is not restarted", async () => {
    const h = live(LiveAdapter);
    await h.manager.ensure(h.session, h.settings, h.context);
    await h.manager.ensure(h.session, { ...h.settings, model: "b", effort: "high" }, h.context);
    expect(h.adapter.starts).toHaveLength(1);
    expect(h.adapter.closes).toEqual([]);
    expect(h.session).toMatchObject({ providerModel: "b", providerEffort: "high" });
  });

  it("still restarts for what only a restart can change, such as the instructions", async () => {
    const h = live(LiveAdapter);
    await h.manager.ensure(h.session, h.settings, h.context);
    await h.manager.ensure(h.session, { ...h.settings, systemPromptRevision: "other" }, h.context);
    expect(h.adapter.starts).toHaveLength(2);
  });

  it("restarts for a model change when the CLI can only change effort live", async () => {
    const h = live(EffortOnlyAdapter);
    await h.manager.ensure(h.session, h.settings, h.context);
    await h.manager.ensure(h.session, { ...h.settings, effort: "high" }, h.context);
    expect(h.adapter.starts).toHaveLength(1);
    await h.manager.ensure(h.session, { ...h.settings, effort: "high", model: "b" }, h.context);
    expect(h.adapter.starts).toHaveLength(2);
  });
});
