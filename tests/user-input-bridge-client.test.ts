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
