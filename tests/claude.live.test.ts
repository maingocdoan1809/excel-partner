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
    expect(toolNames.every((n) => n === "AskUserQuestion" || n.startsWith("mcp__excel__")), JSON.stringify(toolNames)).toBe(true);
    expect(toolNames).toHaveLength(exposedToolNames(false).length + 1);
  }, 120_000);

  it("returns a native AskUserQuestion answer through the pane to the real CLI", async () => {
    const h = await setup([
      { tool: "AskUserQuestion", input: { questions: [{ question: "Where should the result go?", header: "Output",
        options: [{ label: "Current sheet", description: "Use this sheet" }, { label: "New sheet", description: "Create a sheet" }], multiSelect: false }] } },
      { text: "Received your choice." },
    ]);
    const normalHandler = h.pane.onMessage;
    h.pane.onMessage = message => {
      if (message.type === "tool_request" && message.tool === "__request_user_input") {
        h.pane.send({ type: "tool_result", callId: message.callId,
          result: { ok: true, data: { cancelled: false, answers: { "0": ["New sheet"] } } } });
      } else normalHandler?.(message);
    };
    const result = await h.turn("Ask me where to put the result");
    expect(result.finished.reason).toBe("completed");
    expect(h.pane.received.some(m => m.type === "tool_request" && m.tool === "__request_user_input")).toBe(true);
    const last = fake!.requests.filter(r => r.body.tools?.length).at(-1)!;
    expect(JSON.stringify(last.body.messages)).toContain("New sheet");
    expect(result.text).toContain("Received your choice.");
  }, 120_000);

  it("keeps conversation history across turns (resume) ", async () => {
    const h = await setup([{ text: "noted" }]);
    expect((await h.turn("remember ZEBRA")).finished.reason).toBe("completed");
    expect((await h.turn("what did I say?")).finished.reason).toBe("completed");

    const last = fake!.requests.filter((r) => r.body.tools?.length).at(-1)!.body.messages;
    expect(JSON.stringify(last)).toContain("ZEBRA");
  }, 180_000);

  it("passes a pending user message through the next MCP tool result to the real CLI", async () => {
    const h = await setup([{ tool: "get_selection", input: {} }, { text: "Received the update." }]);
    const running = h.turn("Inspect the workbook");
    await h.pane.waitFor(message => message.type === "turn_started");
    h.pane.send({ type: "followup", requestId: "h-1", id: "followup-real-cli", prompt: "Also keep all existing formulas intact." });
    await h.pane.waitFor(message => message.type === "followup_status" && message.id === "followup-real-cli" && message.state === "pending");
    expect((await running).finished.reason).toBe("completed");
    expect(h.pane.received.some(message => message.type === "followup_status" && message.id === "followup-real-cli" && message.state === "delivered")).toBe(true);
    const afterTool = fake!.requests.filter(request => request.body.tools?.length).at(-1)!;
    expect(JSON.stringify(afterTool.body.messages)).toContain("Also keep all existing formulas intact.");
  }, 120_000);

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
