import { execFileSync } from "node:child_process";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { exposedToolNames } from "../protocol/src/index.js";
import { FakeAnthropic, type FakeStep } from "./support/fake-anthropic.js";
import { Harness } from "./support/harness.js";

/**
 * Runs the REAL `claude` binary against a scripted fake Anthropic API. This
 * verifies our flags, MCP wiring, permission handling, stream-json parsing,
 * resume and cancel. It does not verify model behaviour (see scripts/live-smoke).
 */
function hasClaude(): boolean {
  try {
    execFileSync(process.env.CLAUDE_BIN ?? "claude", ["--version"], { stdio: "ignore", timeout: 5000, windowsHide: true });
    return true;
  } catch {
    return false;
  }
}

describe.skipIf(!hasClaude())("claude adapter (real CLI, fake model)", () => {
  let fake: FakeAnthropic | null = null;
  let harness: Harness | null = null;
  const saved = { url: process.env.ANTHROPIC_BASE_URL, key: process.env.ANTHROPIC_API_KEY };

  async function setup(script: FakeStep[]) {
    fake = new FakeAnthropic(script);
    await fake.start();
    process.env.ANTHROPIC_BASE_URL = fake.url;
    process.env.ANTHROPIC_API_KEY = "sk-ant-fake";
    harness = await Harness.start("claude", { allowOfficeJs: false });
    return harness;
  }

  afterEach(async () => {
    await harness?.close();
    await fake?.stop();
    harness = null;
    fake = null;
    if (saved.url === undefined) delete process.env.ANTHROPIC_BASE_URL;
    else process.env.ANTHROPIC_BASE_URL = saved.url;
    if (saved.key === undefined) delete process.env.ANTHROPIC_API_KEY;
    else process.env.ANTHROPIC_API_KEY = saved.key;
  });

  it("runs Excel tools through MCP and streams normalised events", async () => {
    const h = await setup([
      { tool: "get_selection", input: {} },
      { tool: "set_formula", input: { address: "D2", formula: "B2*C2" } },
      { text: "Added the formula." },
    ]);
    const result = await h.turn("add a total");

    expect(result.finished.reason).toBe("completed");
    expect(h.wb.calls.map((c) => c.tool)).toEqual(["get_selection", "set_formula"]);
    expect(h.wb.get("Sheet1", "D2")).toBe("=B2*C2");
    expect(result.text).toBe("Added the formula.");
    const types = result.events.map((e) => e.type);
    expect(types.filter((t) => t === "turn_finished")).toHaveLength(1);
    expect(types).toContain("tool_started");
    expect(result.events.find((e) => e.type === "tool_finished")).toMatchObject({ ok: true });

    // Only the Excel MCP tools are exposed: no Bash/Edit/Read built-ins.
    // Newer CLI versions include a non-callable deferred-tool schema placeholder.
    const toolNames: string[] = fake!.requests.find((r) => r.body.tools?.length)!.body.tools.map((t: any) => t.name)
      .filter((name: string) => name !== "DeferredToolPlaceholder");
    expect(toolNames.every((n) => n.startsWith("mcp__excel__")), JSON.stringify(toolNames)).toBe(true);
    expect(toolNames).toHaveLength(exposedToolNames(false).length);
  }, 120_000);

  it("keeps conversation history across turns (resume) ", async () => {
    const h = await setup([{ text: "noted" }]);
    expect((await h.turn("remember ZEBRA")).finished.reason).toBe("completed");
    expect((await h.turn("what did I say?")).finished.reason).toBe("completed");

    const last = fake!.requests.filter((r) => r.body.tools?.length).at(-1)!.body.messages;
    expect(JSON.stringify(last)).toContain("ZEBRA");
  }, 180_000);

  it("reports tool errors from the pane without crashing the turn", async () => {
    const h = await setup([
      { tool: "read_range", input: { address: "A1", sheet: { name: "Nope" } } },
      { text: "That sheet does not exist." },
    ]);
    const result = await h.turn("read Nope!A1");
    const finished = result.events.find((e) => e.type === "tool_finished");
    expect(finished).toMatchObject({ ok: false });
    expect(result.finished.reason).toBe("completed");
  }, 120_000);

  it("cancels a running turn once, then accepts the next prompt", async () => {
    const h = await setup([{ text: "slow answer", delayMs: 30_000 }]);
    const started = Date.now();
    const cancelled = await h.turn("do something slow", { cancelAfterMs: 4000, timeoutMs: 30_000 });
    expect(cancelled.finished.reason).toBe("cancelled");
    expect(Date.now() - started).toBeLessThan(15_000);
    expect(cancelled.events.filter((e) => e.type === "turn_finished")).toHaveLength(1);

    fake!["script"].splice(0, 1, { text: "fast answer" });
    const next = await h.turn("now answer fast");
    expect(next.finished.reason).toBe("completed");
    expect(next.text).toContain("fast answer");
  }, 120_000);
});
