import { describe, expect, it } from "vitest";
import { executeWithPermission } from "../addin/taskpane/permissions.js";
import type { ToolResult } from "../protocol/src/index.js";

describe("permission mode", () => {
  it("Ask checks occupied cells and retries only after approval", async () => {
    const calls: unknown[] = [];
    const execute = async (_tool: string, args: unknown): Promise<ToolResult> => {
      calls.push(args);
      return (args as { allowOverwrite: boolean }).allowOverwrite
        ? { ok: true, data: { written: 1 } }
        : { ok: false, error: { code: "would_overwrite", message: "Would overwrite A1. Ask the user to confirm." } };
    };
    let asked = "";
    const result = await executeWithPermission("write_range", { address: "A1", allowOverwrite: true }, "ask", execute, async (message) => {
      asked = message;
      return true;
    });
    expect(result.ok).toBe(true);
    expect(calls).toEqual([{ address: "A1", allowOverwrite: false }, { address: "A1", allowOverwrite: true }]);
    expect(asked).toContain("Would overwrite A1");
  });

  it("Ask stops when the user declines sheet deletion", async () => {
    const calls: unknown[] = [];
    const result = await executeWithPermission("manage_sheet", { operation: "delete", confirmDelete: true }, "ask", async (_tool, args) => {
      calls.push(args);
      return { ok: false, error: { code: "confirmation_required", message: "Deleting sheet removes data." } };
    }, async () => false);
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(calls).toEqual([{ operation: "delete", confirmDelete: false }]);
  });

  it("Auto enables overwrite and sheet deletion without asking", async () => {
    const calls: unknown[] = [];
    const execute = async (_tool: string, args: unknown): Promise<ToolResult> => {
      calls.push(args);
      return { ok: true, data: {} };
    };
    const ask = async () => { throw new Error("should not ask"); };
    await executeWithPermission("copy_to", { source: "A1", destination: "B1" }, "auto", execute, ask);
    await executeWithPermission("manage_sheet", { operation: "delete" }, "auto", execute, ask);
    expect(calls).toEqual([
      { source: "A1", destination: "B1", allowOverwrite: true },
      { operation: "delete", confirmDelete: true },
    ]);
  });

  it("leaves unrelated tools and errors alone", async () => {
    const args = { operation: "rename", name: "Sheet1" };
    const result = await executeWithPermission("manage_sheet", args, "auto", async (_tool, received) => {
      expect(received).toBe(args);
      return { ok: false, error: { code: "sheet_not_found", message: "missing" } };
    }, async () => { throw new Error("should not ask"); });
    expect(result).toMatchObject({ ok: false, error: { code: "sheet_not_found" } });
  });
});
