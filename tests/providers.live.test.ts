import { describe, expect, it } from "vitest";
import type { ProviderId } from "../protocol/src/index.js";
import { Harness } from "./support/harness.js";

/**
 * Opt-in tests against REAL, logged-in provider CLIs and real models:
 *
 *   npm run build:server
 *   LIVE_PROVIDERS=codex,copilot npx vitest run tests/providers.live.test.ts
 *
 * They cost model tokens, so they never run by default.
 */
const live = (process.env.LIVE_PROVIDERS ?? "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean) as ProviderId[];

const TASK =
  "Look at my selected row. Put the header 'Total' in D1, the formula =B2*C2 in D2 and =B3*C3 in D3, then bold A1:D1. Use the Excel tools.";

for (const provider of live) {
  describe(`${provider} (real CLI, real model)`, () => {
    it("performs a multi-step Excel task through MCP", async () => {
      const h = await Harness.start(provider);
      try {
        const result = await h.turn(TASK, { timeoutMs: 240_000 });
        expect(result.finished.reason).toBe("completed");
        expect(h.wb.get("Sheet1", "D1")).toBe("Total");
        expect(String(h.wb.get("Sheet1", "D2")).replace(/\s/g, "")).toBe("=B2*C2");
        expect(String(h.wb.get("Sheet1", "D3")).replace(/\s/g, "")).toBe("=B3*C3");
        expect(JSON.stringify(h.wb.formats)).toContain('"bold":true');
        expect(result.events.filter((e) => e.type === "turn_finished")).toHaveLength(1);
        expect(result.events.some((e) => e.type === "tool_started")).toBe(true);
        expect(result.events.filter((e) => e.type === "tool_finished").every((e) => (e as any).ok)).toBe(true);
        expect(result.text.length).toBeGreaterThan(0);
        // No CLI-internal notices leak into the chat stream.
        expect(result.text).not.toMatch(/Disabled tools/);
      } finally {
        await h.close();
      }
    }, 300_000);

    it("keeps context across turns", async () => {
      const h = await Harness.start(provider);
      try {
        await h.turn("For this task my codeword is ZEBRA. Reply only OK, do not call tools.", { timeoutMs: 120_000 });
        const second = await h.turn("What is my codeword? One word, no tools.", { timeoutMs: 120_000 });
        expect(second.text.toUpperCase()).toContain("ZEBRA");
      } finally {
        await h.close();
      }
    }, 300_000);

    it("cancels a turn, finishes exactly once and accepts the next prompt", async () => {
      const h = await Harness.start(provider);
      try {
        const started = Date.now();
        const cancelled = await h.turn(
          "Write a very long, detailed essay (at least 2000 words) about the history of spreadsheets. Do not use tools.",
          { cancelAfterMs: 4000, timeoutMs: 60_000 },
        );
        expect(cancelled.finished.reason).toBe("cancelled");
        expect(cancelled.events.filter((e) => e.type === "turn_finished")).toHaveLength(1);
        expect(Date.now() - started).toBeLessThan(30_000);

        const next = await h.turn("Reply with exactly: READY. No tools.", { timeoutMs: 120_000 });
        expect(next.finished.reason).toBe("completed");
        expect(next.text.toUpperCase()).toContain("READY");
      } finally {
        await h.close();
      }
    }, 300_000);
  });
}

describe.skipIf(live.length < 2)("provider switching", () => {
  it("does not leak history between providers on the same workbook", async () => {
    const [a, b] = live as [ProviderId, ProviderId];
    const h = await Harness.start(a);
    try {
      await h.turn("Remember the secret word PINEAPPLE. Reply only OK, no tools.", { timeoutMs: 120_000 });
      h.pane.send({ type: "set_provider", provider: b });
      await h.pane.waitFor((m, i) => i > 3 && m.type === "provider_status" && m.provider === b, 30_000);
      const reply = await h.turn(
        "What secret word did I ask you to remember? If you do not know, say UNKNOWN. No tools.",
        { timeoutMs: 120_000 },
      );
      expect(reply.text.toUpperCase()).not.toContain("PINEAPPLE");
    } finally {
      await h.close();
    }
  }, 300_000);
});
