import { afterEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { CopilotUsageReader } from "../providers/src/copilot-usage.js";
import { CopilotAdapter } from "../providers/src/copilot.js";
import { CumulativeUsage } from "../providers/src/usage.js";
import { UsageStore } from "../bridge/src/usage.js";
import { formatCredits, usageCsv } from "../addin/taskpane/ui/usage.js";
import type { ProviderEvent } from "../protocol/src/index.js";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function fixture() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "copilot-raw-usage-"));
  dirs.push(dir);
  return path.join(dir, "events.jsonl");
}
const sessionId = "test-cli-session";
// Numeric accounting fixture only: two user requests, twelve underlying model calls.
// Input already includes cache read and cache write; output includes reasoning.
const calls = [
  [30720, 30, 0, 30716, 7709800000],
  [31278, 151, 30716, 560, 598560000],
  [31517, 394, 31276, 239, 766910000],
  [31922, 76, 31515, 405, 492800000],
  [32713, 258, 31920, 789, 775250000],
  [34409, 398, 32709, 1698, 1149990000],
  [34961, 426, 34407, 552, 908470000],
  [35444, 425, 34959, 483, 895740000],
  [36972, 316, 35442, 1528, 1052820000],
  [37303, 105, 36970, 331, 557850000],
  [37823, 126, 37301, 520, 629410000],
  [38715, 739, 37821, 892, 1340610000],
] as const;
function event(index: number) {
  const [inputTokens, outputTokens, cacheReadTokens, cacheWriteTokens, totalNanoAiu] = calls[index]!;
  return { type: "session.usage_record", id: `event-${index}`, data: { usage: {
    inputTokens, outputTokens, cacheReadTokens, cacheWriteTokens,
    accounting: { sourceSessionId: sessionId, usageId: `call-${index}` },
    copilotUsage: { totalNanoAiu }, aiCreditsStatus: "complete",
  } } };
}
const lines = (start: number, end: number) => calls.slice(start, end).map((_, i) => JSON.stringify(event(start + i))).join("\n") + "\n";

describe("original Copilot session accounting", () => {
  it("matches original totals, keeps exact credits/cache split, and ignores shutdown duplicates", async () => {
    const file = fixture();
    fs.writeFileSync(file, lines(0, 4));
    const reader = new CopilotUsageReader(file, sessionId);
    const counter = new CumulativeUsage();
    const first = counter.next((await reader.snapshot())!);
    expect(first).toMatchObject({ inputTokens: 125437, outputTokens: 651, totalTokens: 126088,
      cachedInputTokens: 93507, cacheWriteTokens: 31920, uncachedInputTokens: 10,
      copilotCreditNanoAiu: 9568070000, copilotCredits: 9.56807, source: "copilot_events" });
    fs.appendFileSync(file, lines(4, 12) + JSON.stringify({ type: "session.shutdown", data: { totalNanoAiu: 16878210000 } }) + "\n");
    // Replayed call and another session's accounting must not inflate this session.
    fs.appendFileSync(file, JSON.stringify(event(0)) + "\n");
    const foreign = event(1);
    foreign.data.usage.accounting.sourceSessionId = "another-session";
    fs.appendFileSync(file, JSON.stringify(foreign) + "\n");
    const total = (await reader.snapshot())!;
    expect(total).toMatchObject({ inputTokens: 413777, outputTokens: 3444, totalTokens: 417221,
      cachedInputTokens: 375036, cacheWriteTokens: 38713, uncachedInputTokens: 28,
      copilotCreditNanoAiu: 16878210000, copilotCredits: 16.87821 });
    const second = counter.next(total);
    expect(second).toMatchObject({ totalTokens: 291133, copilotCreditNanoAiu: 7310140000, copilotCredits: 7.31014 });
    expect(counter.next((await reader.snapshot())!)).toMatchObject({ totalTokens: 0, copilotCreditNanoAiu: 0, copilotCredits: 0 });
    const ledgerFile = path.join(path.dirname(file), "usage.jsonl");
    const store = new UsageStore(ledgerFile);
    for (const [i, usage] of [first, second].entries()) store.append({ sessionId: "chat", requestId: String(i),
      provider: "copilot", model: "test-model", startedAt: 1000, durationMs: 100, reason: "completed", usage });
    const summary = new UsageStore(ledgerFile).summaries()[0]!;
    expect(summary).toMatchObject({ totalTokens: 417221, copilotCredits: 16.87821, cacheWriteTokens: 38713,
      uncachedInputTokens: 28, exactCreditReportedTurns: 2, cacheSplitReportedTurns: 2, originalReportedTurns: 2,
      providerSessionIds: [sessionId] });
    const csv = usageCsv([summary]);
    expect(csv).toContain('"16.87821","2","complete (original)"');
    expect(csv).toContain('"28","38713","2","complete","16878210000"');
  });

  it("waits for a complete final line and tolerates a journal that does not exist yet", async () => {
    const file = fixture();
    const reader = new CopilotUsageReader(file, sessionId);
    expect(await reader.snapshot()).toBeNull();
    const line = JSON.stringify(event(0));
    fs.writeFileSync(file, line.slice(0, 40));
    expect(await reader.snapshot()).toBeNull();
    fs.appendFileSync(file, line.slice(40) + "\n");
    expect(await reader.snapshot()).toMatchObject({ totalTokens: 30750, copilotCredits: 7.7098 });
    expect(await reader.snapshot()).toMatchObject({ totalTokens: 30750 });
  });

  it("does not fabricate cache write or credits when the original fields are absent", async () => {
    const file = fixture();
    fs.writeFileSync(file, JSON.stringify({ type: "session.usage_record", id: "legacy", data: { usage: {
      inputTokens: 20, outputTokens: 3, cacheReadTokens: 10,
    } } }) + "\n");
    const usage = (await new CopilotUsageReader(file, sessionId).snapshot())!;
    expect(usage.totalTokens).toBe(23);
    expect(usage.cacheWriteTokens).toBeUndefined();
    expect(usage.uncachedInputTokens).toBeUndefined();
    expect(usage.copilotCredits).toBeUndefined();
  });

  it("uses journal data without ACP usage, including on a cancelled or failed request", async () => {
    for (const fail of [false, true]) {
      const file = fixture();
      fs.writeFileSync(file, lines(0, 4));
      const adapter = new CopilotAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
      const session = { acpSessionId: sessionId, primed: false, active: null,
        usageReader: new CopilotUsageReader(file, sessionId),
        peer: { alive: true, notify() {}, request: async () => {
          if (fail) throw new Error("request failed");
          return { stopReason: "cancelled" };
        } } };
      (adapter as unknown as { sessions: Map<string, unknown> }).sessions.set("provider", session);
      const events: ProviderEvent[] = [];
      for await (const e of adapter.sendTurn({ providerSessionId: "provider", requestId: "request",
        prompt: "hello", selection: null, signal: new AbortController().signal })) events.push(e);
      expect(events.find(e => e.type === "usage_update")).toMatchObject({ usage: {
        source: "copilot_events", totalTokens: 126088, copilotCredits: 9.56807 } });
      expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: fail ? "error" : "cancelled" });
    }
  });

  it("labels a mixed legacy report and preserves exact credit precision in CSV", () => {
    const store = new UsageStore(null);
    const usage = { inputTokens: 2, outputTokens: 1, cachedInputTokens: 0, totalTokens: 3 };
    const base = { sessionId: "chat", provider: "copilot", model: null, startedAt: 1000, durationMs: 10, reason: "completed" as const };
    store.append({ ...base, requestId: "old", usage: { ...usage, copilotCredits: 0.45 } });
    store.append({ ...base, requestId: "new", usage: { ...usage, copilotCredits: 0.000000001, copilotCreditNanoAiu: 1 } });
    const csv = usageCsv(store.summaries());
    expect(csv).toContain('"0.450000001","2","complete (mixed original/rounded)"');
    expect(csv).toContain('"0","unavailable"');
    expect(formatCredits(0)).toBe("0");
    expect(formatCredits(10)).toBe("10");
    expect(formatCredits(0.000000001)).toBe("0.000000001");
  });

  it("uses ACP when the journal is behind and does not double count after it catches up", async () => {
    const file = fixture();
    fs.writeFileSync(file, lines(0, 1));
    const adapter = new CopilotAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
    let turn = 0;
    const session = { acpSessionId: sessionId, primed: false, active: null,
      usageReader: new CopilotUsageReader(file, sessionId),
      peer: { alive: true, notify() {}, request: async () => {
        turn++;
        if (turn > 1) fs.writeFileSync(file, lines(0, 12));
        return { stopReason: "end_turn", usage: turn === 1
          ? { inputTokens: 125437, outputTokens: 651, cachedReadTokens: 93507, totalTokens: 126088 }
          : { inputTokens: 413777, outputTokens: 3444, cachedReadTokens: 375036, totalTokens: 417221 } };
      } } };
    (adapter as unknown as { sessions: Map<string, unknown> }).sessions.set("provider", session);
    const usages = [];
    for (let i = 0; i < 2; i++) {
      for await (const e of adapter.sendTurn({ providerSessionId: "provider", requestId: String(i),
        prompt: "hello", selection: null, signal: new AbortController().signal })) {
        if (e.type === "usage_update") usages.push(e.usage);
      }
    }
    expect(usages[0]).toMatchObject({ source: "copilot_acp", totalTokens: 126088 });
    expect(usages[1]).toMatchObject({ source: "copilot_events", totalTokens: 291133 });
    // Earlier missing cache/credit baselines cannot be reconstructed as this turn's usage.
    expect(usages[1]?.copilotCreditNanoAiu).toBeUndefined();
    expect(usages[1]?.cacheWriteTokens).toBeUndefined();
    expect(usages.reduce((sum, u) => sum + u.totalTokens, 0)).toBe(417221);
  });

  it("rejects invalid original accounting rather than inventing an uncached value", async () => {
    const file = fixture();
    const invalid = { type: "session.usage_record", id: "invalid", data: { usage: {
      inputTokens: 10, outputTokens: 1, cacheReadTokens: 9, cacheWriteTokens: 5,
    } } };
    fs.writeFileSync(file, JSON.stringify(invalid) + "\n");
    await expect(new CopilotUsageReader(file, sessionId).snapshot()).rejects.toThrow("cache exceeds input");
  });
});
