// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { BridgeClient } from "../addin/taskpane/bridge-client.js";
import { showUserInput } from "../addin/taskpane/ui/user-input.js";

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
  const runTool = vi.fn(async () => ({ ok: true as const, data: {} }));
  const client = new BridgeClient({ onState() {}, onHello() {}, onProviderStatus() {}, onProviderEvent() {}, onError() {}, runTool }, "ws://localhost");
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

it("refuses an approved write still queued when stopped", async () => {
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  const runTool = vi.fn(async (_context, tool, args, execute) => { await gate; return execute(tool, args); });
  const client = new BridgeClient({ onState() {}, onHello() {}, onProviderStatus() {}, onProviderEvent() {}, onError() {}, runTool }, "ws://localhost");
  const send = vi.fn();
  (client as any).ws = { readyState: WebSocket.OPEN, send, close: vi.fn() };
  client.setPermissionMode("auto");
  const pending = (client as any).onMessage(JSON.stringify({ type: "tool_request", tool: "clear_range", callId: "write", requestId: "turn", timeoutMs: 120000, args: { address: "A1" } }));
  client.cancel("turn");
  release();
  await pending;
  expect(send.mock.calls.map(([raw]) => JSON.parse(raw))).toContainEqual({ type: "tool_result", callId: "write", result: { ok: false, error: { code: "cancelled", message: "The request ended before this change could run." } } });
});
