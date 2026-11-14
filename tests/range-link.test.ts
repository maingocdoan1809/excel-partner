import { describe, expect, it } from "vitest";
import type { WebSocket } from "ws";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { OfficeRuntime } from "../bridge/src/office-runtime.js";
import type { ToolResult } from "../protocol/src/index.js";

const limits = { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 };

/** A workbook whose pane answers every tool request with `reply`; returns what the assistant gets back. */
function ask(tool: string, args: unknown, reply: ToolResult): Promise<ToolResult> {
  const sessions = new WorkbookSessions({ ttlMs: 60_000 });
  const session = sessions.create("range-link-connection", null);
  session.attachSocket({
    readyState: 1, close() {},
    send(raw: string) {
      const message = JSON.parse(raw) as { type: string; callId?: string };
      if (message.type === "tool_request") queueMicrotask(() => session.resolveTool(message.callId!, reply));
    },
  } as unknown as WebSocket);
  return new OfficeRuntime(sessions, limits, 1000).invoke({ sessionId: session.id, tool, args });
}

const dataOf = (result: ToolResult) => (result as { data: Record<string, unknown> }).data;

describe("a tool result that names a range carries it as a ready-made link", () => {
  it("names the sheet and the range, whatever the tool", async () => {
    const read = await ask("read_range", { sheet: { name: "Sheet2" }, address: "C5" }, { ok: true, data: { sheet: "Sheet2", address: "C5", values: [[1]] } });
    expect(dataOf(read)).toMatchObject({ values: [[1]], link: "[C5](cell:?sheet=Sheet2&range=C5)" });
    const wrote = await ask("write_range", { sheet: { name: "Báo cáo (Q1)" }, address: "A1", values: [["x"]] },
      { ok: true, data: { sheet: "Báo cáo (Q1)", address: "$A$1:$B$2", rowsWritten: 2 } });
    expect(dataOf(wrote).link).toBe("[A1:B2](cell:?sheet=B%C3%A1o%20c%C3%A1o%20(Q1)&range=A1:B2)");
  });

  it("leaves a result alone when it names no sheet, no usable range, or already has a link", async () => {
    for (const data of [
      { address: "C5" },
      { sheet: "", address: "C5" },
      { sheet: "Sheet2" },
      { sheet: "Sheet2", address: "A1:B2,D4" },
      { sheet: "Sheet2", address: "not a range" },
      { sheet: "Sheet2", address: 5 },
    ]) {
      expect(dataOf(await ask("read_range", { address: "A1" }, { ok: true, data })), JSON.stringify(data)).toEqual(data);
    }
    const own = { sheet: "Sheet2", address: "C5", link: "kept" };
    expect(dataOf(await ask("read_range", { address: "A1" }, { ok: true, data: own }))).toEqual(own);
  });

  it("leaves errors and results that are not an object as they are", async () => {
    const failed: ToolResult = { ok: false, error: { code: "range_not_found", message: "no such range" } };
    expect(await ask("read_range", { address: "A1" }, failed)).toEqual(failed);
    expect(await ask("read_range", { address: "A1" }, { ok: true, data: [1, 2] })).toEqual({ ok: true, data: [1, 2] });
    expect(await ask("read_range", { address: "A1" }, { ok: true, data: null })).toEqual({ ok: true, data: null });
  });
});
