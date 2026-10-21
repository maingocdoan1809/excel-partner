import { afterEach, describe, expect, it } from "vitest";
import type { ProviderId } from "../protocol/src/index.js";
import { Harness } from "./support/harness.js";

const BIN_VARS: Record<ProviderId, string> = {
  claude: "CLAUDE_BIN",
  codex: "CODEX_BIN",
  copilot: "COPILOT_BIN",
  opencode: "OPENCODE_BIN",
};

const saved = { ...process.env };
afterEach(() => {
  for (const key of Object.values(BIN_VARS)) {
    if (saved[key] === undefined) delete process.env[key];
    else process.env[key] = saved[key];
  }
});

describe.each(Object.keys(BIN_VARS) as ProviderId[])("%s with a missing CLI", (provider) => {
  it("reports not_installed with an install hint, and fails the turn cleanly", async () => {
    process.env[BIN_VARS[provider]] = "/nonexistent/definitely-not-a-cli";
    const h = await Harness.start(provider);
    try {
      const status = h.pane.received.find((m) => m.type === "provider_status");
      expect(status).toMatchObject({
        detection: { provider, installed: false, status: "not_installed" },
      });
      expect((status as any).detection.hint).toMatch(/Install/);

      const result = await h.turn("hello", { timeoutMs: 30_000 });
      expect(result.finished.reason).toBe("error");
      expect(result.events.filter((e) => e.type === "turn_finished")).toHaveLength(1);
      expect(result.events.find((e) => e.type === "error")).toBeTruthy();

      // The workbook is not left "busy": a second prompt gets the same clean failure, not `busy`.
      const again = await h.turn("hello again", { timeoutMs: 30_000 });
      expect(again.events.find((e) => e.type === "error")).not.toMatchObject({ code: "busy" });
    } finally {
      await h.close();
    }
  }, 60_000);
});
