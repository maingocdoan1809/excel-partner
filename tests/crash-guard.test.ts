import { describe, expect, it } from "vitest";
import { EventEmitter } from "node:events";
import { CrashGuard } from "../bridge/src/crash-guard.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";

describe("CrashGuard", () => {
  it("ignores key order and expires after the block time", () => {
    let now = 0;
    const guard = new CrashGuard(10 * 60_000, () => now);
    guard.record("get_objects", { sheet: { name: "Overview" }, name: undefined });
    expect(guard.blockedMinutes("get_objects", { name: undefined, sheet: { name: "Overview" } })).toBe(10);
    expect(guard.blockedMinutes("get_objects", { sheet: { name: "Other" } })).toBe(0);
    expect(guard.blockedMinutes("get_sheets", { sheet: { name: "Overview" } })).toBe(0);
    now = 10 * 60_000 + 1;
    expect(guard.blockedMinutes("get_objects", { sheet: { name: "Overview" } })).toBe(0);
  });
});

describe("ToolGateway after Excel drops mid-call", () => {
  const limits = { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 };

  function pane() {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const first = sessions.create("conn-1", null);
    const sent: unknown[] = [];
    const socket = Object.assign(new EventEmitter(), { readyState: 1, send: (data: string) => sent.push(JSON.parse(data)) });
    first.attachSocket(socket as never);
    return { sessions, first, socket, sent, gateway: new ToolGateway(sessions, limits, 1000) };
  }

  it("refuses the identical call in the session Excel opens after restarting, but not a different one", async () => {
    const { sessions, first, socket, sent, gateway } = pane();
    const crashing = gateway.invoke({ sessionId: first.id, tool: "get_objects", args: { sheet: { name: "Overview" } } });
    await new Promise(resolve => setTimeout(resolve, 10));
    expect(sent).toHaveLength(1);
    first.detachSocket(socket as never); // Excel died while the call was running
    expect(await crashing).toMatchObject({ ok: false, error: { code: "disconnected" } });

    const second = sessions.create("conn-2", null);
    const retry = await gateway.invoke({ sessionId: second.id, tool: "get_objects", args: { sheet: { name: "Overview" } } });
    expect(retry).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect((retry as { error: { message: string } }).error.message).toMatch(/most likely crashed/);

    // A different call is not blocked; with no pane attached it fails for the usual reason.
    const other = await gateway.invoke({ sessionId: second.id, tool: "get_objects", args: { sheet: { name: "Cover" } } });
    expect(other).toMatchObject({ ok: false, error: { code: "disconnected", message: "task pane is not connected" } });
  });

  it("does not block a call that simply failed", async () => {
    const { first, sent, gateway } = pane();
    const call = gateway.invoke({ sessionId: first.id, tool: "get_sheets", args: {} });
    await new Promise(resolve => setTimeout(resolve, 10));
    const request = sent[0] as { callId: string };
    first.resolveTool(request.callId, { ok: false, error: { code: "office_error", message: "boom" } });
    await call;
    const again = gateway.invoke({ sessionId: first.id, tool: "get_sheets", args: {} });
    await new Promise(resolve => setTimeout(resolve, 10));
    expect(sent).toHaveLength(2);
    first.resolveTool((sent[1] as { callId: string }).callId, { ok: true, data: [] });
    expect(await again).toMatchObject({ ok: true });
  });
});
