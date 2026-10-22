import { describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderEvent, ProviderId, UsageSummary } from "../protocol/src/index.js";
import { PaneClient } from "./helpers.js";

class NativeAdapter extends MockAdapter {
  starts: StartSessionContext[] = [];
  prompts: string[] = [];
  native = new Map<string, string>();
  count = 0;
  override async startSession(ctx: StartSessionContext): Promise<string> {
    this.starts.push(ctx);
    if (ctx.resumeSessionId === "missing") throw new Error("Session not found");
    const id = await super.startSession(ctx);
    this.native.set(id, ctx.resumeSessionId ?? `native-${++this.count}`);
    return id;
  }
  async sessionTitle(id: string): Promise<string> { return `Native title ${id}`; }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    this.prompts.push(input.prompt);
    const native = this.native.get(input.providerSessionId)!;
    yield { type: "session_info", provider: this.id, providerSessionId: native, title: `Native title ${native}`, requestId: input.requestId };
    yield { type: "turn_started", requestId: input.requestId };
    yield { type: "usage_update", requestId: input.requestId, usage: { inputTokens: 100, outputTokens: 20,
      totalTokens: 120, cachedInputTokens: 30, providerSessionId: native } };
    yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
  }
}

async function connect(url: string, provider: ProviderId, connectionId: string) {
  const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: connectionId });
  const hello = await pane.waitFor(m => m.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("Handshake failed");
  pane.send({ type: "set_provider", provider });
  await pane.waitFor(m => m.type === "provider_status");
  return { pane, headers: { authorization: `Bearer ${hello.sessionToken}` } };
}
async function prompt(pane: PaneClient, requestId: string, chatSessionId: string, resumeSessionId?: string) {
  pane.send({ type: "prompt", requestId, prompt: "new request", selection: null, chatSessionId,
    resumeSessionId, context: resumeSessionId ? "DO_NOT_REPLAY_NATIVE_HISTORY" : undefined });
  await pane.waitFor(m => m.type === "turn_finished" && m.requestId === requestId);
}

describe("native session identity through bridge, ledger and reconnect", () => {
  it.each(["copilot", "claude", "codex", "opencode"] as const)("%s resumes the same accounting row after reconnect and creates a new row only for New chat", async provider => {
    const adapter = new NativeAdapter({ id: provider });
    const bridge = new Bridge(loadConfig({ port: 0, tls: null, logLevel: "error" }), {
      usageFile: null, registry: new ProviderRegistry().register(adapter),
    });
    const { url } = await bridge.listen();
    let active = await connect(url, provider, "first-connection");
    try {
      await prompt(active.pane, "first", "provisional-chat");
      await active.pane.close();
      active = await connect(url, provider, "second-connection");
      await prompt(active.pane, "resumed", `${provider}:native-1`, "native-1");
      let data = await (await fetch(`${url}/session/usage`, { headers: active.headers })).json() as { sessions: UsageSummary[] };
      expect(data.sessions).toHaveLength(1);
      expect(data.sessions[0]).toMatchObject({ sessionId: `${provider}:native-1`, turns: 2, totalTokens: 240,
        title: "Native title native-1", providerSessionIds: ["native-1"] });
      expect(adapter.starts[1]?.resumeSessionId).toBe("native-1");
      expect(adapter.prompts[1]).not.toContain("DO_NOT_REPLAY_NATIVE_HISTORY");
      await prompt(active.pane, "fresh", "new-chat-button");
      data = await (await fetch(`${url}/session/usage`, { headers: active.headers })).json() as { sessions: UsageSummary[] };
      expect(data.sessions.map(s => s.sessionId).sort()).toEqual([`${provider}:native-1`, `${provider}:native-2`]);
      expect(adapter.starts[2]?.resumeSessionId).toBeUndefined();
      await prompt(active.pane, "missing-resume", `${provider}:missing`, "missing");
      expect(active.pane.received.find(m => m.type === "error" && m.requestId === "missing-resume"))
        .toMatchObject({ code: "resume_failed", message: expect.stringContaining("No replacement session was created") });
      expect(adapter.count).toBe(2);
    } finally { await active.pane.close(); await bridge.close(); }
  });

  it("keeps the same native ID from two providers in separate namespaces", async () => {
    const copilot = new NativeAdapter({ id: "copilot" });
    const codex = new NativeAdapter({ id: "codex" });
    const bridge = new Bridge(loadConfig({ port: 0, tls: null, logLevel: "error" }), {
      usageFile: null, registry: new ProviderRegistry().register(copilot).register(codex),
    });
    const { url } = await bridge.listen();
    const active = await connect(url, "copilot", "test-connection");
    try {
      await prompt(active.pane, "copilot-request", "copilot-chat");
      active.pane.send({ type: "set_provider", provider: "codex" });
      await active.pane.waitFor(m => m.type === "provider_status" && m.provider === "codex");
      await prompt(active.pane, "codex-request", "codex-chat");
      const data = await (await fetch(`${url}/session/usage`, { headers: active.headers })).json() as { sessions: UsageSummary[] };
      expect(data.sessions.map(s => s.sessionId).sort()).toEqual(["codex:native-1", "copilot:native-1"]);
    } finally { await active.pane.close(); await bridge.close(); }
  });
});
