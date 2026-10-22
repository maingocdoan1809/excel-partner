import { describe, expect, it } from "vitest";
import { SelectionSnapshotSchema } from "../protocol/src/index.js";
import { AsyncQueue } from "../providers/src/async-queue.js";
import { compareVersions, extractVersion } from "../providers/src/base.js";
import { ClaudeStreamTranslator } from "../providers/src/claude.js";
import { JsonRpcPeer } from "../providers/src/jsonrpc.js";
import { buildTurnPrompt } from "../providers/src/prompt.js";

async function collect<T>(iterable: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of iterable) out.push(item);
  return out;
}

describe("AsyncQueue", () => {
  it("drains buffered items then completes", async () => {
    const q = new AsyncQueue<number>();
    q.push(1);
    q.push(2);
    q.end();
    q.push(3); // ignored after end
    expect(await collect(q)).toEqual([1, 2]);
  });

  it("wakes a waiting consumer and propagates failure after draining", async () => {
    const q = new AsyncQueue<number>();
    const consumer = collect(q);
    setTimeout(() => {
      q.push(1);
      q.fail(new Error("boom"));
    }, 10);
    await expect(consumer).rejects.toThrow("boom");
  });
});

describe("version helpers", () => {
  it("extracts and compares versions", () => {
    expect(extractVersion("codex-cli 0.162.1")).toBe("0.162.1");
    expect(extractVersion("2.1.294 (Claude Code)")).toBe("2.1.294");
    expect(extractVersion("GitHub Copilot CLI 1.0.95.")).toBe("1.0.95");
    expect(compareVersions("0.162.1", "0.162.1")).toBe(0);
    expect(compareVersions("0.9.0", "0.162.1")).toBe(-1);
    expect(compareVersions("1.0.0", "0.162.1")).toBe(1);
  });
});

describe("buildTurnPrompt", () => {
  it("passes the prompt through without a selection", () => {
    expect(buildTurnPrompt("hello", null)).toBe("hello");
  });

  it("keeps the selected shape through snapshot validation and prompt serialization", () => {
    const selection = SelectionSnapshotSchema.parse({
      workbookName: "flow.xlsx", sheet: "S", sheetId: null, address: "", rowCount: 0, columnCount: 0,
      values: [], headers: null, truncated: false, shapeSelectionStatus: "active",
      activeShape: { id: "s1", name: "Decision", type: "FlowChartDecision", text: "Approved?", left: 10, top: 20, width: 60, height: 30 },
    });
    const json = JSON.parse(buildTurnPrompt("validate this flow", selection).split("\n")[1]!);
    expect(json.activeShape).toEqual(selection.activeShape);
    expect(json.shapeSelectionStatus).toBe("active");
  });

  it("embeds the selection snapshot and truncates huge selections", () => {
    const values = Array.from({ length: 1000 }, (_, r) => [r, "x", "y"]);
    const text = buildTurnPrompt("sum it", {
      workbookName: "a.xlsx",
      sheet: "S",
      sheetId: null,
      address: "A1:C1000",
      rowCount: 1000,
      columnCount: 3,
      values,
      headers: null,
      truncated: false,
    });
    expect(text).toContain("<selection_snapshot>");
    expect(text.endsWith("sum it")).toBe(true);
    const json = JSON.parse(text.split("\n")[1]!);
    expect(json.values.length).toBeLessThan(1000);
    expect(json.truncated).toBe(true);
  });
});

describe("ClaudeStreamTranslator", () => {
  it("maps text deltas, tool use/result and completion", () => {
    const t = new ClaudeStreamTranslator("r1");
    const events = [
      { type: "system", subtype: "init" },
      { type: "stream_event", event: { type: "message_start", message: { id: "m1" } } },
      { type: "stream_event", event: { type: "content_block_delta", delta: { type: "text_delta", text: "Hel" } } },
      { type: "assistant", message: { id: "m1", content: [{ type: "text", text: "Hello" }] } },
      {
        type: "assistant",
        message: { id: "m2", content: [{ type: "tool_use", id: "t1", name: "mcp__excel__read_range", input: { address: "A1" } }] },
      },
      { type: "user", message: { content: [{ type: "tool_result", tool_use_id: "t1", content: "boom", is_error: true }] } },
      { type: "result", is_error: false, subtype: "success" },
    ].flatMap((m) => t.translate(m));

    expect(events.map((e) => e.type)).toEqual([
      "message_delta",
      "message_completed",
      "tool_started",
      "tool_finished",
      "turn_finished",
    ]);
    expect(events[2]).toMatchObject({ tool: "read_range", callId: "t1", args: { address: "A1" } });
    expect(events[3]).toMatchObject({ ok: false, tool: "read_range", error: "boom" });
    expect(events[4]).toMatchObject({ reason: "completed" });
  });

  it("emits a delta for assistant text that was not streamed", () => {
    const t = new ClaudeStreamTranslator("r1");
    const events = t.translate({ type: "assistant", message: { id: "m9", content: [{ type: "text", text: "Hi" }] } });
    expect(events.map((e) => e.type)).toEqual(["message_delta", "message_completed"]);
  });

  it("turns an API failure into login_required + error finish and ignores sub-agent traffic", () => {
    const t = new ClaudeStreamTranslator("r1");
    expect(t.translate({ type: "assistant", parent_tool_use_id: "x", message: { content: [{ type: "text", text: "no" }] } })).toEqual([]);
    const events = t.translate({ type: "result", is_error: true, api_error_status: 403, result: "org disabled" });
    expect(events[0]).toMatchObject({ type: "error", code: "login_required", message: "org disabled" });
    expect(events[1]).toMatchObject({ type: "turn_finished", reason: "error" });
  });
});

describe("JsonRpcPeer", () => {
  const echoServer = `
    const rl = require("readline").createInterface({ input: process.stdin });
    rl.on("line", (l) => {
      const m = JSON.parse(l);
      if (m.method === "echo") process.stdout.write(JSON.stringify({ id: m.id, result: m.params }) + "\\n");
      else if (m.method === "fail") process.stdout.write(JSON.stringify({ id: m.id, error: { code: -1, message: "nope" } }) + "\\n");
      else if (m.method === "ask") {
        process.stdout.write(JSON.stringify({ id: 99, method: "client/question", params: { q: 1 } }) + "\\n");
        process.stdout.write(JSON.stringify({ method: "note", params: { n: 1 } }) + "\\n");
        process.stdout.write(JSON.stringify({ id: m.id, result: "asked" }) + "\\n");
      } else if (m.method === "die") process.exit(3);
    });
  `;

  function spawnPeer(extra: Partial<ConstructorParameters<typeof JsonRpcPeer>[0]> = {}) {
    return new JsonRpcPeer({ command: process.execPath, args: ["-e", echoServer], ...extra });
  }

  it("correlates responses, errors and notifications", async () => {
    const notes: unknown[] = [];
    const peer = spawnPeer({
      onNotification: (m, p) => notes.push([m, p]),
      onRequest: (m, p) => ({ answered: m, p }),
    });
    try {
      expect(await peer.request("echo", { a: 1 })).toEqual({ a: 1 });
      await expect(peer.request("fail")).rejects.toThrow("nope");
      expect(await peer.request("ask")).toBe("asked");
      expect(notes).toEqual([["note", { n: 1 }]]);
    } finally {
      await peer.close();
    }
  });

  it("rejects pending requests when the process dies and refuses new ones", async () => {
    const peer = spawnPeer();
    const pending = peer.request("die");
    await expect(pending).rejects.toThrow(/exited/);
    await expect(peer.request("echo", {})).rejects.toThrow(/not running/);
    expect(peer.alive).toBe(false);
  });

  it("times out requests the peer never answers", async () => {
    const peer = spawnPeer();
    try {
      await expect(peer.request("silent", {}, 100)).rejects.toThrow(/timed out/);
    } finally {
      await peer.close();
    }
  });
});
