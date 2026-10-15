import { describe, expect, it } from "vitest";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";

const limits = { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 };

function setup() {
  const sessions = new SessionManager({ ttlMs: 60_000 });
  const session = sessions.create("conn-1", null);
  const gateway = new ToolGateway(sessions, limits, 1000);
  return { sessions, session, gateway };
}

describe("ToolGateway", () => {
  it("rejects an unknown session", async () => {
    const { gateway } = setup();
    const result = await gateway.invoke({ sessionId: "missing", tool: "get_sheets", args: {} });
    expect(result).toMatchObject({ ok: false, error: { code: "disconnected" } });
  });

  it("rejects an unknown tool", async () => {
    const { gateway, session } = setup();
    const result = await gateway.invoke({ sessionId: session.id, tool: "rm_rf", args: {} });
    expect(result).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
  });

  it("rejects invalid arguments", async () => {
    const { gateway, session } = setup();
    const result = await gateway.invoke({
      sessionId: session.id,
      tool: "read_range",
      args: { address: "1A" },
    });
    expect(result).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("enforces the read size limit", async () => {
    const { gateway, session } = setup();
    const result = await gateway.invoke({
      sessionId: session.id,
      tool: "read_range",
      args: { address: "A1:Z1000" },
    });
    expect(result).toMatchObject({ ok: false, error: { code: "size_limit_exceeded" } });
  });

  it("reports disconnected when the pane is not attached", async () => {
    const { gateway, session } = setup();
    const result = await gateway.invoke({
      sessionId: session.id,
      tool: "read_range",
      args: { address: "A1" },
    });
    expect(result).toMatchObject({ ok: false, error: { code: "disconnected" } });
  });
});
