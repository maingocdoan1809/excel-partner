import { describe, expect, it } from "vitest";
import { ClaudeAdapter, ClaudeStreamTranslator } from "../providers/src/claude/adapter.js";
import type { StartSessionContext } from "../providers/src/types.js";
import { CodexAdapter } from "../providers/src/codex/adapter.js";
import { CopilotAdapter } from "../providers/src/copilot/adapter.js";
import { OpencodeAdapter } from "../providers/src/opencode/adapter.js";
import type { ProviderEvent } from "../protocol/src/index.js";

const runtime = { bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } };
const thoughts = (events: ProviderEvent[]) => events.flatMap(event => event.type === "thinking_delta" ? [event.text] : []);

describe("Claude thinking", () => {
  const delta = (thinking: string) => ({ type: "stream_event", event: { type: "content_block_delta", delta: { type: "thinking_delta", thinking } } });
  const start = (type: string) => ({ type: "stream_event", event: { type: "content_block_start", content_block: { type } } });

  it("streams thinking deltas and separates one block from the next", () => {
    const t = new ClaudeStreamTranslator("r1");
    const events = [
      { type: "stream_event", event: { type: "message_start", message: { id: "m1" } } },
      start("thinking"), delta("Check the "), delta("header row."),
      start("thinking"), delta("Then sort."),
      // The complete message repeats what already streamed.
      { type: "assistant", message: { id: "m1", content: [{ type: "thinking", thinking: "Check the header row.Then sort." }] } },
    ].flatMap(message => t.translate(message));
    expect(thoughts(events)).toEqual(["Check the ", "header row.", "\n\nThen sort."]);
  });

  it("shows thinking that only arrives in the complete message", () => {
    const t = new ClaudeStreamTranslator("r1");
    const events = t.translate({ type: "assistant", message: { id: "m2", content: [
      { type: "thinking", thinking: "Plan first." }, { type: "thinking", thinking: "Then write." }, { type: "text", text: "Done" }] } });
    expect(events.map(event => event.type)).toEqual(["thinking_delta", "thinking_delta", "message_delta", "message_completed"]);
    expect(thoughts(events)).toEqual(["Plan first.", "\n\nThen write."]);
  });

  it("says nothing when the model returns no readable thinking", () => {
    const t = new ClaudeStreamTranslator("r1");
    expect(t.translate(start("thinking"))).toEqual([]);
    expect(t.translate({ type: "assistant", message: { id: "m3", content: [{ type: "thinking", thinking: "" }] } })).toEqual([]);
  });

  it("ignores a sub-agent's thinking", () => {
    const t = new ClaudeStreamTranslator("r1");
    expect(t.translate({ ...delta("secret"), parent_tool_use_id: "agent" })).toEqual([]);
  });
});

describe("Claude thinking-display flag", () => {
  const context = {} as StartSessionContext;
  function adapter(open: (withFlag: boolean) => Promise<string>) {
    const claude = new ClaudeAdapter(runtime);
    const calls: boolean[] = [];
    (claude as unknown as { openSession: (ctx: unknown, withFlag: boolean) => Promise<string> }).openSession = async (_ctx, withFlag) => { calls.push(withFlag); return open(withFlag); };
    return { claude, calls };
  }

  it("asks for summaries first", async () => {
    const { claude, calls } = adapter(async () => "id");
    expect(await claude.startSession(context)).toBe("id");
    expect(calls).toEqual([true]);
  });

  it("opens again without the flag when the CLI does not know it, and remembers for later sessions", async () => {
    const { claude, calls } = adapter(async withFlag => {
      if (withFlag) throw new Error("claude exited (code=1, signal=null): error: unknown option '--thinking-display'");
      return "id";
    });
    expect(await claude.startSession(context)).toBe("id");
    expect(await claude.startSession(context)).toBe("id");
    expect(calls).toEqual([true, false, false]);
  });

  it("does not hide other failures behind a retry", async () => {
    const { claude, calls } = adapter(async () => { throw new Error("Session missing"); });
    await expect(claude.startSession(context)).rejects.toThrow("Session missing");
    expect(calls).toEqual([true]);
  });
});

describe("Claude replay-user-messages flag", () => {
  it("opens again without it when the CLI does not know it, keeps the thinking flag, and remembers", async () => {
    const claude = new ClaudeAdapter(runtime);
    const calls: Array<[boolean, boolean]> = [];
    (claude as unknown as { openSession: (ctx: unknown, thinking: boolean, replay: boolean) => Promise<string> }).openSession = async (_ctx, thinking, replay) => {
      calls.push([thinking, replay]);
      if (replay) throw new Error("claude exited (code=1, signal=null): error: unknown option '--replay-user-messages'");
      return "id";
    };
    expect(await claude.startSession({} as StartSessionContext)).toBe("id");
    expect(await claude.startSession({} as StartSessionContext)).toBe("id");
    expect(calls).toEqual([[true, true], [true, false], [true, false]]);
  });
});

describe("Codex reasoning", () => {
  function run(notifications: Array<[string, Record<string, unknown>]>) {
    const events: ProviderEvent[] = [];
    const session = { threadId: "thread", active: { queue: { push: (event: ProviderEvent) => events.push(event) }, requestId: "r", turnId: "1", interrupting: false } };
    const adapter = new CodexAdapter(runtime) as unknown as { onNotification(session: unknown, method: string, params: unknown): void };
    for (const [method, params] of notifications) adapter.onNotification(session, method, { threadId: "thread", ...params });
    return events;
  }

  it("shows the summary, one paragraph per part, and not the raw text next to it", () => {
    const events = run([
      ["item/reasoning/summaryTextDelta", { itemId: "i1", summaryIndex: 0, delta: "Read the " }],
      ["item/reasoning/summaryTextDelta", { itemId: "i1", summaryIndex: 0, delta: "sheet." }],
      ["item/reasoning/summaryTextDelta", { itemId: "i1", summaryIndex: 1, delta: "Write totals." }],
      ["item/reasoning/textDelta", { itemId: "i1", contentIndex: 0, delta: "raw reasoning" }],
    ]);
    expect(thoughts(events)).toEqual(["Read the ", "sheet.", "\n\nWrite totals."]);
  });

  it("shows the raw reasoning of models that send no summary", () => {
    const events = run([
      ["item/reasoning/textDelta", { itemId: "i1", contentIndex: 0, delta: "Look at A1." }],
      ["item/reasoning/textDelta", { itemId: "i2", contentIndex: 0, delta: "Now B1." }],
    ]);
    expect(thoughts(events)).toEqual(["Look at A1.", "\n\nNow B1."]);
  });
});

describe("ACP thought chunks", () => {
  const chunk = (text: string, type = "text") => ({ sessionId: "native", update: { sessionUpdate: "agent_thought_chunk", content: { type, text } } });

  it.each([
    ["Copilot", () => new CopilotAdapter(runtime)],
    ["OpenCode", () => new OpencodeAdapter(runtime)],
  ])("%s turns agent_thought_chunk into thinking_delta", (_name, create) => {
    const events: ProviderEvent[] = [];
    const session = { acpSessionId: "native", active: { queue: { push: (event: ProviderEvent) => events.push(event) }, requestId: "r", text: "", tools: new Map() } };
    const adapter = create() as unknown as { onNotification(session: unknown, method: string, params: unknown): void };
    adapter.onNotification(session, "session/update", chunk("Consider the totals."));
    adapter.onNotification(session, "session/update", chunk(""));
    adapter.onNotification(session, "session/update", chunk("image", "image"));
    expect(events).toEqual([{ type: "thinking_delta", requestId: "r", text: "Consider the totals." }]);
  });
});
