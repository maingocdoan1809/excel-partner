import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ProviderEvent } from "../protocol/src/index.js";

const rpc = vi.hoisted(() => ({ options: null as any, calls: [] as Array<{ method: string; params: any }> }));
vi.mock("../providers/src/jsonrpc.js", () => ({ JsonRpcPeer: class {
  alive = true; stderrTail = [];
  constructor(options: any) { rpc.options = options; }
  async request(method: string, params: any) {
    rpc.calls.push({ method, params });
    if (method === "thread/start") return { thread: { id: "native-thread" } };
    if (method === "thread/read") return { thread: { id: "native-thread" } };
    if (method === "skills/list") return { data: [{ skills: [{ name: "report", description: "Build a report", enabled: true, path: "/skills/report/SKILL.md" }] }] };
    if (["turn/start", "review/start", "thread/compact/start"].includes(method)) {
      setImmediate(() => rpc.options.onNotification("turn/completed", { threadId: "native-thread", turn: { id: "turn", status: "completed" } }));
      return { turn: { id: "turn", status: "inProgress" } };
    }
    return {};
  }
  notify() {} onExit() {} async close() {}
} }));
import { CodexAdapter } from "../providers/src/codex.js";
let adapter: CodexAdapter, id: string;
beforeEach(async () => {
  rpc.calls = [];
  adapter = new CodexAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
  id = await adapter.startSession({ workbookConnectionId: "workbook", sessionId: "bridge", signal: new AbortController().signal,
    mcp: { command: "node", args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" } });
});
afterEach(async () => { await adapter.dispose(); });
async function command(name: string, args = "") {
  const events: ProviderEvent[] = [];
  for await (const event of adapter.sendTurn({ providerSessionId: id, requestId: "cmd", prompt: "WORKBOOK CONTEXT MUST NOT BE USED", selection: null,
    command: { name, args }, signal: new AbortController().signal })) events.push(event);
  return events;
}
describe("Codex command API mapping", () => {
  it("lists skills without starting model inference", async () => {
    const events = await command("skills");
    expect(events).toContainEqual(expect.objectContaining({ type: "message_completed", text: expect.stringContaining("report") }));
    expect(rpc.calls.some(call => call.method === "turn/start")).toBe(false);
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "completed" });
  });
  it("injects a selected skill as a native skill input", async () => {
    await command("skills", "report make a summary");
    expect(rpc.calls.find(call => call.method === "turn/start")!.params.input).toEqual([
      { type: "text", text: "$report make a summary", text_elements: [] }, { type: "skill", name: "report", path: "/skills/report/SKILL.md" },
    ]);
  });
  it.each([["compact", "thread/compact/start"], ["review", "review/start"]])("maps /%s and waits for native completion", async (name, method) => {
    const events = await command(name!);
    expect(rpc.calls.find(call => call.method === method)!.params.threadId).toBe("native-thread");
    expect(events.filter(event => event.type === "turn_finished")).toHaveLength(1);
    expect(rpc.calls.some(call => call.method === "turn/start")).toBe(false);
  });
});
