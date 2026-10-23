import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { StartSessionContext } from "../providers/src/types.js";
import type { UserInputResponse } from "../protocol/src/index.js";

const rpc = vi.hoisted(() => ({ options: null as any }));
vi.mock("../providers/src/jsonrpc.js", () => ({ JsonRpcPeer: class {
  alive = true;
  stderrTail = [];
  constructor(options: any) { rpc.options = options; }
  async request(method: string) {
    if (method === "thread/start") return { thread: { id: "native-thread" } };
    if (method === "thread/read") return { thread: {} };
    if (method === "turn/start") return { turn: { id: "native-turn", status: "inProgress" } };
    return {};
  }
  notify() {}
  onExit() {}
  async close() {}
} }));
import { CodexAdapter } from "../providers/src/codex.js";
let adapter: CodexAdapter;
beforeEach(() => { adapter = new CodexAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } }); });
afterEach(async () => { await adapter.dispose(); });

describe("Codex native input server requests", () => {
  it.each(["reply", "resolved"])("routes the real app-server method and handles %s", async action => {
    let signal: AbortSignal | undefined;
    const ask: StartSessionContext["requestUserInput"] = vi.fn(async (_request, requestSignal): Promise<UserInputResponse> => {
      signal = requestSignal;
      if (action === "resolved") await new Promise<void>(resolve => requestSignal!.addEventListener("abort", () => resolve(), { once: true }));
      return { cancelled: action === "resolved", answers: action === "resolved" ? {} : { location: ["New sheet"] } };
    });
    const id = await adapter.startSession({ requestUserInput: ask, workbookConnectionId: "workbook", sessionId: "bridge", signal: new AbortController().signal,
      mcp: { command: "node", args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" } });
    const events = adapter.sendTurn({ providerSessionId: id, requestId: "r", prompt: "Ask where", selection: null, signal: new AbortController().signal })[Symbol.asyncIterator]();
    await events.next();
    await events.next();
    const next = events.next();
    await Promise.resolve();
    const pending = rpc.options.onRequest("item/tool/requestUserInput", { threadId: "native-thread", turnId: "native-turn", itemId: "item",
      questions: [{ id: "location", header: "Output", question: "Where?", options: [{ label: "New sheet", description: "Create one" }] }] }, 42);
    if (action === "resolved") rpc.options.onNotification("serverRequest/resolved", { threadId: "native-thread", requestId: 42 });
    expect(await pending).toEqual({ answers: { location: { answers: action === "resolved" ? [] : ["New sheet"] } } });
    expect(ask).toHaveBeenCalledOnce();
    if (action === "resolved") expect(signal?.aborted).toBe(true);
    expect(rpc.options.onRequest("item/tool/requestUserInput", { threadId: "foreign" }, 43)).toEqual({ answers: {} });
    rpc.options.onNotification("turn/completed", { threadId: "native-thread", turn: { id: "native-turn", status: "completed" } });
    await next;
    for await (const _event of { [Symbol.asyncIterator]: () => events }) { /* drain cleanup */ }
  });
});
