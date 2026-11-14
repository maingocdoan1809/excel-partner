// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { BridgeClient } from "../frontend/taskpane/bridge-client.js";
import { hostAdapterFor } from "../frontend/taskpane/host/index.js";
import { showUserInput } from "../frontend/taskpane/ui/user-input.js";

describe("pane native question lifecycle", () => {
  it.each(["stop", "finish", "provider", "disconnect"])("dismisses a pending question on %s even in Auto mode", async reason => {
    const send = vi.fn();
    const ask = vi.fn((request, signal) => showUserInput(document.body, request, signal));
    const client = new BridgeClient({ onUserInput: ask, onState() {}, onHello() {}, onProviderStatus() {},
      onProviderEvent() {}, onError(message) { throw new Error(message); } }, "ws://localhost");
    const socket = { readyState: WebSocket.OPEN, send, close: vi.fn() };
    (client as any).ws = socket;
    client.setPermissionMode("auto");
    const pending = (client as any).onMessage(JSON.stringify({ type: "tool_request", tool: "__request_user_input",
      callId: "q", requestId: "turn", timeoutMs: 600_000, args: { questions: [{ id: "answer", question: "Where?" }] } }));
    expect(ask).toHaveBeenCalledOnce();
    expect(document.querySelector("form")).not.toBeNull();
    expect(send).not.toHaveBeenCalled();
    if (reason === "stop") client.cancel("turn");
    else if (reason === "finish") await (client as any).onMessage(JSON.stringify({ type: "turn_finished", requestId: "turn", reason: "completed" }));
    else if (reason === "provider") await (client as any).onMessage(JSON.stringify({ type: "user_input_closed", callId: "q" }));
    else client.disconnect();
    await pending;
    expect(document.querySelector("form")).toBeNull();
    if (reason !== "disconnect") expect(send.mock.calls.map(([raw]) => JSON.parse(raw))).toContainEqual({
      type: "tool_result", callId: "q", result: { ok: true, data: { cancelled: true, answers: {} } },
    });
  });
});

it.each(["stop", "finish", "disconnect"])("dismisses write approval on %s", async reason => {
  const runTool = vi.spyOn(hostAdapterFor("excel"), "execute").mockResolvedValue({ ok: true, data: {} });
  const client = new BridgeClient({ onState() {}, onHello() {}, onProviderStatus() {}, onProviderEvent() {}, onError() {} }, "ws://localhost");
  (client as any).ws = { readyState: WebSocket.OPEN, send: vi.fn(), close: vi.fn() };
  const pending = (client as any).onMessage(JSON.stringify({ type: "tool_request", tool: "execute_office_js", callId: "write", requestId: "turn", timeoutMs: 120000, args: { code: "return 1" } }));
  expect(document.querySelector(".permission-overlay")).not.toBeNull();
  if (reason === "stop") client.cancel("turn");
  else if (reason === "finish") await (client as any).onMessage(JSON.stringify({ type: "turn_finished", requestId: "turn", reason: "cancelled" }));
  else client.disconnect();
  await pending;
  expect(document.querySelector(".permission-overlay")).toBeNull();
  expect(runTool).not.toHaveBeenCalled();
});


it("runs file changes one at a time and refuses a queued change once the request is stopped", async () => {
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  const order: string[] = [];
  const host = hostAdapterFor("excel");
  vi.spyOn(host, "execute").mockImplementation(async (tool: string) => {
    order.push(`start ${tool}`);
    if (tool === "write_range") await gate;
    order.push(`end ${tool}`);
    return { ok: true, data: {} };
  });
  const client = new BridgeClient({ onState() {}, onHello() {}, onProviderStatus() {}, onProviderEvent() {}, onError() {} }, "ws://localhost");
  const send = vi.fn();
  (client as any).ws = { readyState: WebSocket.OPEN, send, close: vi.fn() };
  client.setPermissionMode("auto");
  const request = (tool: string, callId: string, requestId: string, args: unknown) => (client as any).onMessage(JSON.stringify({ type: "tool_request", tool, callId, requestId, timeoutMs: 120000, args }));
  const first = request("write_range", "w1", "turn-a", { address: "A1", values: [[1]] });
  const second = request("clear_range", "w2", "turn-b", { address: "A1" });
  const third = request("clear_range", "w3", "turn", { address: "B1" });
  const read = request("read_range", "r1", "turn-a", { address: "A1" });
  await read;
  expect(order).toContain("end read_range");
  expect(order).not.toContain("start clear_range");
  expect(order).not.toContain("end write_range");
  client.cancel("turn");
  release();
  await Promise.all([first, second, third]);
  const starts = order.filter((entry) => entry.startsWith("start"));
  expect([...starts].sort()).toEqual(["start read_range", "start write_range"]);
  const results = send.mock.calls.map(([raw]) => JSON.parse(raw)).filter((m) => m.type === "tool_result");
  for (const id of ["w2", "w3"]) expect(results.find((m) => m.callId === id).result).toMatchObject({ ok: false, error: { code: "cancelled" } });
});
