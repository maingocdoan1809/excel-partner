import { describe, expect, it, vi } from "vitest";
import { OfficeRuntime } from "../bridge/src/office-runtime.js";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { BATCHABLE_TOOLS, BatchInputSchema, MAX_BATCH_OPERATIONS, SERVER_TOOLS, isToolNameForHost, type ToolResult } from "../protocol/src/index.js";
import { officeToolCatalog } from "../shared/src/catalog.js";

const limits = { maxReadCells: 100_000, maxWriteCells: 100_000, maxSheets: 10, maxSelectionCells: 1000 };

function setup(host: "excel" | "word" = "excel") {
  const sessions = new WorkbookSessions({ ttlMs: 60_000 });
  const session = sessions.create("conn-batch", { host });
  const gateway = new OfficeRuntime(sessions, limits, 1000, undefined, true);
  const requestTool = vi.spyOn(session, "requestTool");
  return { session, gateway, requestTool };
}
const ok = (data: unknown = {}): ToolResult => ({ ok: true, data });
const write = { tool: "write_range", input: { address: "A1", values: [[1]] } };
const read = { tool: "read_range", input: { address: "A1" } };
const clear = { tool: "clear_range", input: { address: "A1" } };

describe("batch", () => {
  it("runs the operations in order through the normal path and returns each result", async () => {
    const { session, gateway, requestTool } = setup();
    requestTool.mockImplementation(async (tool) => ok({ ran: tool }));
    const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write, read, clear] } });
    expect(requestTool.mock.calls.map((call) => call[0])).toEqual(["write_range", "read_range", "clear_range"]);
    expect(result).toEqual({ ok: true, data: { total: 3, completed: 3, results: [
      { tool: "write_range", ok: true, data: { ran: "write_range" } }, { tool: "read_range", ok: true, data: { ran: "read_range" } }, { tool: "clear_range", ok: true, data: { ran: "clear_range" } },
    ] } });
  });

  it("applies the defaults of each tool's own schema, as a single call would", async () => {
    const { session, gateway, requestTool } = setup();
    requestTool.mockResolvedValue(ok());
    await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write] } });
    expect(requestTool.mock.calls[0]![1]).toMatchObject({ address: "A1", allowOverwrite: false });
  });

  it("checks every operation before running any, and says which are wrong", async () => {
    const { session, gateway, requestTool } = setup();
    const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write, { tool: "read_range", input: { address: 12 } }, { tool: "set_formula", input: {} }] } });
    expect(requestTool).not.toHaveBeenCalled();
    expect(result).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("2 of 3 operations are invalid. Nothing was changed.") } });
    const detail = (result as { error: { detail: Array<{ operation: number; tool: string }> } }).error.detail;
    expect(detail.map((entry) => [entry.operation, entry.tool])).toEqual([[2, "read_range"], [3, "set_formula"]]);
  });

  it("stops at the first failure and says how far it got, without rolling back", async () => {
    const { session, gateway, requestTool } = setup();
    requestTool.mockResolvedValueOnce(ok()).mockResolvedValueOnce({ ok: false, error: { code: "would_overwrite", message: "B2 is not empty." } });
    const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write, read, clear] } });
    expect(requestTool).toHaveBeenCalledTimes(2);
    expect(result).toMatchObject({ ok: false, error: { code: "would_overwrite", detail: { total: 3, completed: 1, notRun: 1 } } });
    const message = (result as { error: { message: string } }).error.message;
    expect(message).toContain("Operation 2 of 3 (read_range) failed: B2 is not empty.");
    expect(message).toContain("1 operation(s) were applied and are not rolled back; 1 were not run.");
  });

  it("keeps going after a failure when asked to, and still reports the first failure", async () => {
    const { session, gateway, requestTool } = setup();
    requestTool.mockResolvedValueOnce({ ok: false, error: { code: "range_not_found", message: "No A1." } }).mockResolvedValue(ok());
    const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write, read, clear], stopOnError: false } });
    expect(requestTool).toHaveBeenCalledTimes(3);
    expect(result).toMatchObject({ ok: false, error: { code: "range_not_found", detail: { total: 3, completed: 2, notRun: 0 } } });
  });

  it("does not run when the request has already ended", async () => {
    const { session, gateway, requestTool } = setup();
    const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write] }, signal: AbortSignal.abort() });
    expect(requestTool).not.toHaveBeenCalled();
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
  });

  it("counts each operation in the turn's tool statistics, as well as the batch", async () => {
    const { session, gateway, requestTool } = setup();
    requestTool.mockResolvedValue(ok({ rows: 1 }));
    session.beginTurn("turn-1");
    await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write, read, read] } });
    const stats = session.takeToolStats("turn-1")!;
    expect(stats.batch).toMatchObject({ calls: 1, errors: 0 });
    expect(stats.write_range).toMatchObject({ calls: 1 });
    expect(stats.read_range).toMatchObject({ calls: 2 });
  });
});

describe("what a batch may contain", () => {
  it("refuses tools outside the list, a batch inside a batch, and more than the limit", async () => {
    const { session, gateway, requestTool } = setup();
    for (const operations of [
      [{ tool: "execute_office_js", input: { code: "return 1" } }],
      [{ tool: "batch", input: { operations: [write] } }],
      [{ tool: "insert_image", input: {} }],
      [{ tool: "run_python", input: {} }],
      Array.from({ length: MAX_BATCH_OPERATIONS + 1 }, () => read),
      [],
    ]) {
      const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations } });
      expect(result, JSON.stringify(operations).slice(0, 60)).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    }
    expect(requestTool).not.toHaveBeenCalled();
  });

  it("is a tool of Excel only, answered by the bridge, and listed with the tools it can run", () => {
    expect(isToolNameForHost("excel", "batch")).toBe(true);
    expect(isToolNameForHost("word", "batch")).toBe(false);
    expect(isToolNameForHost("powerpoint", "batch")).toBe(false);
    expect(SERVER_TOOLS).toContain("batch");
    const tool = officeToolCatalog("excel", true, {}).find((candidate) => candidate.name === "batch")!;
    expect(JSON.stringify(tool.inputSchema)).toContain("write_range");
    expect(JSON.stringify(tool.inputSchema)).not.toContain("execute_office_js");
  });

  it("only offers tools that exist for Excel", () => {
    for (const name of BATCHABLE_TOOLS) expect(isToolNameForHost("excel", name), name).toBe(true);
    expect(BatchInputSchema.parse({ operations: [read] }).stopOnError).toBe(true);
  });

  it("answers an Excel tool in a Word session as unavailable, not by running it", async () => {
    const { session, gateway, requestTool } = setup("word");
    const result = await gateway.invoke({ sessionId: session.id, tool: "batch", args: { operations: [write] } });
    expect(result).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
    expect(requestTool).not.toHaveBeenCalled();
  });
});
