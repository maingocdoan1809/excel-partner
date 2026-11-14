import { afterEach, beforeEach, expect, it, vi } from "vitest";

/** Codex: a message sent during a turn is added to it with `turn/steer`, which only works on the turn that is running. */

const rpc = vi.hoisted(() => ({ requests: [] as Array<{ method: string; params: any }>, steerFails: false }));
vi.mock("../providers/src/jsonrpc.js", () => ({ JsonRpcPeer: class {
  alive = true;
  stderrTail = [];
  async request(method: string, params: any) {
    rpc.requests.push({ method, params });
    if (method === "thread/start") return { thread: { id: "native-thread" } };
    if (method === "thread/read") return { thread: {} };
    if (method === "turn/start") return { turn: { id: "native-turn", status: "inProgress" } };
    if (method === "turn/steer") { if (rpc.steerFails) throw new Error("expected turn is not active"); return { turnId: "native-turn" }; }
    return {};
  }
  notify() {}
  onExit() {}
  async close() {}
} }));
import { CodexAdapter } from "../providers/src/codex/adapter.js";

let adapter: CodexAdapter;
beforeEach(() => { rpc.requests = []; rpc.steerFails = false; adapter = new CodexAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } }); });
afterEach(async () => { await adapter.dispose(); });

async function openSession() {
  return adapter.startSession({ requestUserInput: vi.fn(), workbookConnectionId: "workbook", sessionId: "bridge", signal: new AbortController().signal,
    mcp: { command: "node", args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" } });
}

async function runningTurn(id: string, signal = new AbortController().signal) {
  const events = adapter.sendTurn({ providerSessionId: id, requestId: "r", prompt: "Make a report", selection: null, signal })[Symbol.asyncIterator]();
  await events.next();
  await events.next();
  void events.next(); // the turn now waits for the CLI's notifications
  await vi.waitFor(() => expect(rpc.requests.some(request => request.method === "turn/start")).toBe(true));
  await new Promise(resolve => setImmediate(resolve)); // the response, and with it the turn's id, is handled just after
  return events;
}

it("steers the running turn by its id", async () => {
  const id = await openSession();
  await runningTurn(id);
  expect(await adapter.steerTurn(id, "Also use blue")).toBe(true);
  expect(rpc.requests.find(request => request.method === "turn/steer")?.params).toEqual({
    threadId: "native-thread", expectedTurnId: "native-turn", input: [{ type: "text", text: "Also use blue", text_elements: [] }] });
});

it("says no when the CLI refuses, so the caller queues the message", async () => {
  const id = await openSession();
  await runningTurn(id);
  rpc.steerFails = true;
  expect(await adapter.steerTurn(id, "Also use blue")).toBe(false);
});

it("says no when no turn is running", async () => {
  const id = await openSession();
  expect(await adapter.steerTurn(id, "Nothing running")).toBe(false);
  expect(rpc.requests.some(request => request.method === "turn/steer")).toBe(false);
});
