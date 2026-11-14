import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { UsageStore, type UsageRecord } from "../bridge/src/usage.js";

const total = (n: number) => ({ inputTokens: n, outputTokens: n / 10, cachedInputTokens: n / 2, totalTokens: n + n / 10,
  cacheWriteTokens: 0, uncachedInputTokens: n / 2, reasoningTokens: 0, source: "opencode_export" as const, providerSessionId: "ses_1" });
const turn = (startedAt: number, usage: UsageRecord["usage"] = null): UsageRecord => ({
  sessionId: "opencode:ses_1", providerSessionId: "ses_1", requestId: `r${startedAt}`, provider: "opencode", model: null,
  startedAt, durationMs: 100, reason: "completed", usage });

let folder = "";
afterEach(() => { if (folder) fs.rmSync(folder, { recursive: true, force: true }); folder = ""; });

describe("usage that a CLI reports only when asked", () => {
  it("asks for turns that reported nothing, once, and not again until another turn ends", () => {
    const store = new UsageStore(null);
    store.append(turn(1000)); store.append(turn(2000));
    expect(store.unsettled()).toEqual([{ provider: "opencode", providerSessionId: "ses_1" }]);
    store.settle("opencode", "ses_1", total(1000), 3000);
    expect(store.unsettled()).toEqual([]);
    store.append(turn(4000));
    expect(store.unsettled()).toHaveLength(1);
  });

  it("counts the CLI's total for the turns that reported nothing", () => {
    const store = new UsageStore(null);
    store.append(turn(1000)); store.append(turn(2000));
    expect(store.summaries()[0]).toMatchObject({ turns: 2, reportedTurns: 0, totalTokens: 0 });
    store.settle("opencode", "ses_1", total(1000), 3000);
    expect(store.summaries()[0]).toMatchObject({ turns: 2, reportedTurns: 2, inputTokens: 1000, outputTokens: 100, totalTokens: 1100, originalReportedTurns: 2 });
  });

  it("does not count twice what turns of the same conversation already reported", () => {
    const store = new UsageStore(null);
    store.append(turn(1000, { inputTokens: 400, outputTokens: 40, cachedInputTokens: 200, totalTokens: 440 }));
    store.append(turn(2000));
    store.settle("opencode", "ses_1", total(1000), 3000);
    expect(store.summaries()[0]).toMatchObject({ turns: 2, reportedTurns: 2, inputTokens: 1000, totalTokens: 1100 });
  });

  it("leaves a turn that ended after the total as not reported", () => {
    const store = new UsageStore(null);
    store.append(turn(1000));
    store.settle("opencode", "ses_1", total(500), 3000);
    store.append(turn(4000));
    expect(store.summaries()[0]).toMatchObject({ turns: 2, reportedTurns: 1, inputTokens: 500 });
  });

  it("keeps the totals across restarts, and an older program ignores them as damaged lines", () => {
    folder = fs.mkdtempSync(path.join(os.tmpdir(), "usage-settle-"));
    const file = path.join(folder, "usage.jsonl");
    const first = new UsageStore(file);
    first.append(turn(1000));
    first.settle("opencode", "ses_1", total(1000), 3000);
    const second = new UsageStore(file);
    expect(second.unsettled()).toEqual([]);
    expect(second.summaries()[0]).toMatchObject({ turns: 1, reportedTurns: 1, inputTokens: 1000 });
  });
});
