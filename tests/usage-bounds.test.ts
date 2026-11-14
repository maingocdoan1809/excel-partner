import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { USAGE_FILE_MAX_BYTES, USAGE_MAX_AGE_MS, USAGE_MAX_RECORDS, UsageStore, type UsageRecord } from "../bridge/src/usage.js";

let dir = "";
afterEach(() => { if (dir) fs.rmSync(dir, { recursive: true, force: true }); });

const NOW = 1_800_000_000_000;
const record = (index: number, startedAt = NOW - 1000): UsageRecord => ({
  sessionId: `chat-${index % 50}`, requestId: `r${index}`, provider: "claude", model: null, startedAt, durationMs: 10, reason: "completed",
  usage: { inputTokens: 1, outputTokens: 1, cachedInputTokens: 0, totalTokens: 2 } as UsageRecord["usage"],
});

describe("usage ledger bounds", () => {
  it("drops turns older than the panel shows and keeps at most the newest records in memory", () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "usage-bounds-"));
    const file = path.join(dir, "usage.jsonl");
    fs.writeFileSync(file, [record(0, NOW - USAGE_MAX_AGE_MS - 1), ...Array.from({ length: USAGE_MAX_RECORDS + 10 }, (_, i) => record(i + 1))].map(r => JSON.stringify(r)).join("\n"));
    const store = new UsageStore(file);
    const turns = store.summaries().reduce((sum, summary) => sum + summary.turns, 0);
    expect(turns).toBe(USAGE_MAX_RECORDS);
  });

  it("rotates a large ledger into .1 and starts a compact file that still restores the kept turns", () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "usage-bounds-"));
    const file = path.join(dir, "usage.jsonl");
    fs.writeFileSync(file, "x".repeat(USAGE_FILE_MAX_BYTES + 1) + "\n");
    const store = new UsageStore(file);
    store.append(record(1));
    expect(fs.existsSync(`${file}.1`)).toBe(true);
    expect(fs.statSync(file).size).toBeLessThan(10_000);
    expect(new UsageStore(file).summaries()).toEqual(store.summaries());
    expect(store.summaries()[0]).toMatchObject({ sessionId: "chat-1", turns: 1 });
  });
});
