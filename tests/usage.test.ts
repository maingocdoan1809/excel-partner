import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { UsageStore, type UsageRecord } from "../bridge/src/usage.js";
import { usageCsv, formatDuration } from "../frontend/taskpane/ui/usage.js";
import { parseCopilotCredits } from "../providers/src/usage.js";
import { CopilotAdapter } from "../providers/src/copilot/adapter.js";
import { CodexAdapter } from "../providers/src/codex/adapter.js";
import { ClaudeStreamTranslator } from "../providers/src/claude/adapter.js";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput } from "../providers/src/types.js";
import { toolResultChars, type ProviderEvent } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const temporaryFile = () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-usage-"));
  dirs.push(dir);
  return path.join(dir, "usage.jsonl");
};
const tokens = { inputTokens: 100, outputTokens: 20, cachedInputTokens: 30, totalTokens: 120 };
const record = (sessionId: string, requestId: string, usage: UsageRecord["usage"] = tokens): UsageRecord => ({
  sessionId, requestId, provider: "claude", model: null, startedAt: 1000, durationMs: 1500, reason: "completed", usage,
});

describe("usage accounting", () => {
  it("restores per-session totals, includes unknown/cancelled turns, and never adds cache twice", () => {
    const file = temporaryFile();
    const store = new UsageStore(file);
    store.append(record("a", "1", { ...tokens, copilotCredits: 0.45 }));
    store.append({ ...record("a", "2", null), reason: "cancelled" });
    store.append(record("b", "3"));
    const restored = new UsageStore(file).summaries();
    expect(restored.find(s => s.sessionId === "a")).toMatchObject({
      turns: 2, reportedTurns: 1, creditReportedTurns: 1, copilotCredits: 0.45, durationMs: 3000, ...tokens,
    });
    expect(restored.find(s => s.sessionId === "b")?.turns).toBe(1);
    fs.appendFileSync(file, '{"incomplete":');
    const recovered = new UsageStore(file);
    recovered.append(record("b", "4"));
    expect(new UsageStore(file).summaries().find(s => s.sessionId === "b")?.turns).toBe(2);
  });

  it("exports Excel-compatible CSV with blanks for missing usage and escaped titles", () => {
    const store = new UsageStore(null);
    store.append(record("a", "1"));
    store.append(record("a", "2", null));
    store.append(record("b", "3", null));
    const csv = usageCsv(store.summaries(), new Map([["a", '=SUM(1,2) "test"\nnext']]));
    expect(csv.startsWith("\uFEFF")).toBe(true);
    expect(csv).toContain('"\'=SUM(1,2) ""test""\nnext"');
    expect(csv).toContain('"partial"');
    expect(csv).toContain('"","","","","unavailable"');
    expect(formatDuration(3_661_999)).toBe("1h 1m 1s");
    expect(formatDuration(61_000)).toBe("1m 1s");
  });

  it("takes Claude result usage once, including cache reads and writes, even on errors", () => {
    const translator = new ClaudeStreamTranslator("turn");
    const usage = { input_tokens: 100, output_tokens: 20, cache_read_input_tokens: 30, cache_creation_input_tokens: 40 };
    expect(translator.translate({ type: "assistant", message: { usage, content: [] } })).toEqual([]);
    expect(translator.translate({ type: "result", usage, is_error: true })).toMatchObject([
      { type: "usage_update", usage: { inputTokens: 170, outputTokens: 20, cachedInputTokens: 30, totalTokens: 190 } },
      { type: "error" }, { type: "turn_finished", reason: "error" },
    ]);
    expect(translator.translate({ type: "result" })).toEqual([{ type: "turn_finished", requestId: "turn", reason: "completed" }]);
  });

  it("regresses the real Copilot 39,998 overcount using consecutive cumulative prompt responses", async () => {
    const adapter = new CopilotAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
    const responses = [
      { inputTokens: 13280, outputTokens: 14, cachedReadTokens: 10305, totalTokens: 13294 },
      { inputTokens: 26676, outputTokens: 28, cachedReadTokens: 23582, totalTokens: 26704 },
      { inputTokens: 26676, outputTokens: 28, cachedReadTokens: 23582, totalTokens: 26704 },
    ];
    const session = {
      peer: { alive: true, request: async () => ({ stopReason: "end_turn", usage: responses.shift() }), notify() {} },
      acpSessionId: "copilot-cli-session", primed: false, active: null,
    };
    const internal = adapter as unknown as { sessions: Map<string, unknown> };
    internal.sessions.set("provider-session", session);
    const store = new UsageStore(null);
    const deltas = [];
    for (let i = 0; i < 3; i++) {
      const events: ProviderEvent[] = [];
      for await (const event of adapter.sendTurn({ providerSessionId: "provider-session", requestId: `turn-${i}`,
        prompt: "hello", selection: null, signal: new AbortController().signal })) events.push(event);
      const usage = events.find(event => event.type === "usage_update");
      expect(usage?.type).toBe("usage_update");
      if (usage?.type !== "usage_update") throw new Error("missing usage");
      deltas.push(usage.usage);
      store.append(record("chat", `turn-${i}`, usage.usage));
    }
    expect(deltas.map(usage => usage.totalTokens)).toEqual([13294, 13410, 0]);
    expect(deltas[1]).toMatchObject({ inputTokens: 13396, outputTokens: 14, cachedInputTokens: 13277, totalTokens: 13410 });
    expect(store.summaries()[0]).toMatchObject({ totalTokens: 26704, inputTokens: 26676, outputTokens: 28, cachedInputTokens: 23582 });
    // A new CLI session starts its own baseline, even when the pane continues an existing chat.
    internal.sessions.set("new-provider-session", { acpSessionId: "new-cli-session", primed: false, active: null, peer: { ...session.peer,
      request: async () => ({ stopReason: "end_turn", usage: { inputTokens: 100, outputTokens: 10, cachedReadTokens: 0, totalTokens: 110 } }) } });
    let restartedUsage;
    for await (const event of adapter.sendTurn({ providerSessionId: "new-provider-session", requestId: "restart",
      prompt: "hello", selection: null, signal: new AbortController().signal })) {
      if (event.type === "usage_update") restartedUsage = event.usage;
    }
    expect(restartedUsage?.totalTokens).toBe(110);
  });

  it("does not substitute rounded /usage credits when the original journal is unavailable", async () => {
    const adapter = new CopilotAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
    const internal = adapter as unknown as { sessions: Map<string, unknown> };
    const prompts: string[] = [];
    const session = { acpSessionId: "billing-session", primed: false, active: null, supportsUsageCommand: true,
      peer: { alive: true, notify() {}, request: async (_method: string, params: { prompt: Array<{ text: string }> }) => {
        prompts.push(params.prompt[0]!.text);
        return { stopReason: "end_turn", usage: { inputTokens: 100, outputTokens: 10, totalTokens: 110 } };
      } },
    };
    internal.sessions.set("billing", session);
    const events: ProviderEvent[] = [];
    for await (const event of adapter.sendTurn({ providerSessionId: "billing", requestId: "request",
      prompt: "hello", selection: null, signal: new AbortController().signal })) events.push(event);
    const usage = events.find(event => event.type === "usage_update");
    expect(usage).toMatchObject({ usage: { totalTokens: 110, source: "copilot_acp", providerSessionId: "billing-session" } });
    if (usage?.type !== "usage_update") throw new Error("missing usage");
    expect(usage.usage.copilotCredits).toBeUndefined();
    expect(prompts).toHaveLength(1);
    expect(prompts.some(prompt => prompt === "/usage")).toBe(false);
    // Historical rounded credit parsing remains supported for old records.
    expect(parseCopilotCredits("Requests: 0 AI Credits (1s)")).toBe(0);
    expect(parseCopilotCredits("Requests: 3 premium requests")).toBeNull();
    expect(parseCopilotCredits("Requests: <0.01 AI Credits")).toBeNull();
  });

  it("uses Codex thread totals as an odometer across updates and turns", () => {
    const events: ProviderEvent[] = [];
    const queue = { push: (event: ProviderEvent) => events.push(event) };
    const session = { threadId: "thread", active: { queue, requestId: "first", turnId: "1", interrupting: false } };
    const adapter = new CodexAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
    const handler = adapter as unknown as { onNotification(session: unknown, method: string, params: unknown): void };
    const report = (input: number, output: number, cached: number, turnId: string) => handler.onNotification(session, "thread/tokenUsage/updated", {
      threadId: "thread", turnId, tokenUsage: { total: { inputTokens: input, outputTokens: output, cachedInputTokens: cached, totalTokens: input + output } },
    });
    report(100, 20, 30, "1");
    report(200, 50, 60, "1");
    report(200, 50, 60, "1");
    expect(events.at(-1)).toMatchObject({ usage: { inputTokens: 200, outputTokens: 50, totalTokens: 250 } });
    session.active = { queue, requestId: "second", turnId: "2", interrupting: false };
    report(200, 50, 60, "1"); // late notification for the previous turn
    expect(events).toHaveLength(3);
    report(350, 80, 100, "2");
    report(450, 100, 130, "2");
    expect(events.at(-1)).toMatchObject({ requestId: "second", usage: { inputTokens: 250, outputTokens: 50, cachedInputTokens: 70, totalTokens: 300 } });
  });

  it("persists only the last usage snapshot per turn, authenticates access and restores after restart", async () => {
    class UsageAdapter extends MockAdapter {
      override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
        yield { type: "turn_started", requestId: input.requestId };
        yield { type: "usage_update", requestId: input.requestId, usage: tokens };
        yield { type: "usage_update", requestId: input.requestId, usage: { ...tokens, outputTokens: 40, totalTokens: 140 } };
        yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
        yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
      }
    }
    const file = temporaryFile();
    const config = loadConfig({ port: 0, tls: null, logLevel: "error" });
    const start = async () => {
      const bridge = new Bridge(config, { usageFile: file, registry: new ProviderRegistry().register(new UsageAdapter({ id: "claude" })) });
      const { url } = await bridge.listen();
      const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
      pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "usage-test-workbook" });
      const hello = await pane.waitFor(m => m.type === "hello_ok");
      if (hello.type !== "hello_ok") throw new Error("handshake failed");
      return { bridge, url, pane, headers: { authorization: `Bearer ${hello.sessionToken}` } };
    };
    let active = await start();
    try {
      expect((await fetch(`${active.url}/session/usage`)).status).toBe(401);
      active.pane.send({ type: "set_provider", provider: "claude" });
      await active.pane.waitFor(m => m.type === "provider_status");
      for (const [requestId, chatSessionId] of [["1", "chat-a"], ["2", "chat-a"], ["3", "chat-b"]]) {
        active.pane.send({ type: "prompt", requestId: requestId!, chatSessionId, prompt: "hello", selection: null });
        await active.pane.waitFor(m => m.type === "turn_finished" && m.requestId === requestId);
      }
      const response = await fetch(`${active.url}/session/usage?sessionId=chat-a`, { headers: active.headers });
      expect(await response.json()).toMatchObject({ sessions: [{ sessionId: "chat-a", turns: 2, reportedTurns: 2, totalTokens: 280 }] });
      await active.pane.close();
      await active.bridge.close();
      active = await start();
      const restored = await fetch(`${active.url}/session/usage`, { headers: active.headers });
      expect(((await restored.json()) as { sessions: unknown[] }).sessions).toHaveLength(2);
      expect(new UsageStore(file).summaries().find(s => s.sessionId === "chat-a")?.turns).toBe(2);
    } finally {
      await active.pane.close();
      await active.bridge.close();
    }
  });
  it("keeps per-tool calls, errors and result sizes with the turn, and the size of the catalog offered", () => {
    const file = temporaryFile();
    const tools = { read_range: { calls: 3, errors: 1, resultChars: 4200 }, get_sheets: { calls: 1, errors: 0, resultChars: 310 } };
    new UsageStore(file).append({ ...record("a", "1"), tools, catalogChars: 63_000 });
    const line = fs.readFileSync(file, "utf8").split("\n").filter(Boolean).map(entry => JSON.parse(entry)).find(entry => entry.requestId === "1");
    expect(line).toMatchObject({ tools, catalogChars: 63_000 });
    // A ledger written before these fields existed still loads.
    fs.appendFileSync(file, "\n" + JSON.stringify(record("a", "2")) + "\n");
    expect(new UsageStore(file).summaries().find(s => s.sessionId === "a")?.turns).toBe(2);
  });

  it("counts a result by its text, not by a picture's bytes", () => {
    expect(toolResultChars({ ok: true, data: { rows: 2 } })).toBe(JSON.stringify({ rows: 2 }).length);
    expect(toolResultChars({ ok: true, data: { rows: 2, image: { mimeType: "image/png", data: "A".repeat(50_000) } } })).toBe(JSON.stringify({ rows: 2 }).length);
    expect(toolResultChars({ ok: false, error: { code: "x", message: "y" } })).toBe(JSON.stringify({ code: "x", message: "y" }).length);
  });

  it("records which tools a turn used and how much they returned, without arguments", async () => {
    const file = temporaryFile();
    const port = await getFreePort();
    const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
    const bridge = new Bridge(config, { usageFile: file, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", toolCalls: [{ tool: "get_selection", args: {} }, { tool: "get_selection", args: {} }] })) });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    try {
      pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "usage-tools-workbook" });
      await pane.waitFor(m => m.type === "hello_ok");
      pane.send({ type: "set_provider", provider: "claude" });
      await pane.waitFor(m => m.type === "provider_status");
      pane.onMessage = message => {
        if (message.type === "tool_request") pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { address: "A1", secret: "cell contents" } } });
      };
      pane.send({ type: "prompt", requestId: "tools-1", chatSessionId: "chat-tools", prompt: "look", selection: null });
      await pane.waitFor(m => m.type === "turn_finished" && m.requestId === "tools-1");
      const text = fs.readFileSync(file, "utf8");
      const entry = text.split("\n").filter(Boolean).map(line => JSON.parse(line)).find(line => line.requestId === "tools-1");
      expect(entry.tools.get_selection).toMatchObject({ calls: 2, errors: 0, resultChars: 2 * JSON.stringify({ address: "A1", secret: "cell contents" }).length });
      expect(entry.catalogChars).toBeGreaterThan(1000);
      expect(text).not.toContain("cell contents");
    } finally {
      await pane.close();
      await bridge.close();
    }
  });
});
