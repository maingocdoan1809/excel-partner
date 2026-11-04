import { beforeEach, describe, expect, it, vi } from "vitest";
import type { ToolResult } from "../protocol/src/index.js";

vi.mock("../addin/taskpane/executor/photo.js", async (importOriginal) => {
  const real = await importOriginal<typeof import("../addin/taskpane/executor/photo.js")>();
  return { ...real, photograph: vi.fn(), fingerprintNow: vi.fn(), restorePhoto: vi.fn(), locateInsert: vi.fn(), removeInsertedRows: vi.fn() };
});

import * as photoModule from "../addin/taskpane/executor/photo.js";
import { JOURNAL_LIMITS, UndoJournal, describeRestore, isReply } from "../addin/taskpane/executor/journal.js";
import { planUndo } from "../addin/taskpane/executor/undo-plan.js";
import { WRITE_TOOLS } from "../addin/taskpane/executor/index.js";

const { canonicalStyle, cellHash, fingerprintOf, needsTextPrefix, sameFingerprint, TooLarge } = photoModule;
const mocked = {
  photograph: vi.mocked(photoModule.photograph),
  fingerprintNow: vi.mocked(photoModule.fingerprintNow),
  restorePhoto: vi.mocked(photoModule.restorePhoto),
  locateInsert: vi.mocked(photoModule.locateInsert),
  removeInsertedRows: vi.mocked(photoModule.removeInsertedRows),
};

const sheet = { name: "Data" };

describe("planUndo", () => {
  it("records write_range from the anchor cell, grown to the size of the values", () => {
    expect(planUndo("write_range", { sheet, address: "B2", values: [[1, 2, 3], [4, 5]] })).toEqual({
      kind: "cells", style: false, targets: [{ sheet, address: "B2", grow: { rows: 2, columns: 3 } }],
    });
    expect(planUndo("write_range", { address: "A1", values: [[1]], notes: { A1: "n" } })).toMatchObject({ caveats: ["cell notes are not restored"] });
    expect(planUndo("write_range", { address: "A1", values: [] })).toEqual({ kind: "none" });
  });

  it("only photographs styles for tools that change them", () => {
    expect(planUndo("set_formula", { sheet, address: "C2:C9", formula: "=A2" })).toMatchObject({ kind: "cells", style: false });
    expect(planUndo("clear_range", { sheet, address: "A1:B2" })).toMatchObject({ style: false });
    expect(planUndo("clear_range", { sheet, address: "A1:B2", clearType: "formats" })).toMatchObject({ style: true, merged: true, caveats: [expect.stringContaining("conditional formats")] });
    expect(planUndo("clear_range", { sheet, address: "A1:B2", clearType: "all" })).toMatchObject({ style: true, caveats: [expect.stringContaining("data validation")] });
    expect(planUndo("format_range", { sheet, address: "A1:B2", format: { bold: true } })).toMatchObject({ style: true });
    expect(planUndo("sort_range", { sheet, address: "A1:C9", keys: [{ column: "A" }] })).toMatchObject({ style: true });
    expect(planUndo("copy_to", { sheet, source: "A1", destination: "A2:A9", copyType: "values" })).toMatchObject({ style: false });
    expect(planUndo("copy_to", { sheet, source: "A1", destination: "A2:A9" })).toMatchObject({ style: true, merged: true, targets: [{ sheet, address: "A2:A9" }], caveats: [expect.stringContaining("copied from the source")] });
    expect(planUndo("copy_to", { sheet, source: "A1", destination: "A2:A9", copyType: "values" })).not.toHaveProperty("caveats");
  });

  it("captures column widths and row heights only when the tool can change them", () => {
    expect(planUndo("format_range", { sheet, address: "B2:C3", format: { columnWidth: 20 } })).toMatchObject({ layout: { address: "B2:C3", dimensions: ["width"] } });
    expect(planUndo("format_range", { sheet, address: "B2:C3", format: { rowHeight: 30, autofitColumns: true } })).toMatchObject({ layout: { dimensions: ["width", "height"] } });
    expect(planUndo("format_range", { sheet, address: "B2:C3", format: { bold: true } })).not.toHaveProperty("layout");
    expect(planUndo("resize_range", { sheet, address: "A:C", width: { type: "points", value: 20 } })).toEqual({
      kind: "cells", style: false, targets: [], layout: { sheet, address: "A:C", dimensions: ["width"] },
    });
    expect(planUndo("resize_range", { sheet, width: { type: "points", value: 20 } })).toMatchObject({ kind: "irreversible" });
    expect(planUndo("resize_range", { sheet, address: "A:C" })).toEqual({ kind: "none" });
  });

  it("find_replace is a no-op preview unless it applies, and covers every sheet when none is named", () => {
    expect(planUndo("find_replace", { find: "a", replace: "b" })).toEqual({ kind: "none" });
    expect(planUndo("find_replace", { find: "a", replace: "b", apply: true })).toMatchObject({ targets: [{ allSheets: true }] });
    expect(planUndo("find_replace", { sheet, find: "a", replace: "b", apply: true })).toMatchObject({ targets: [{ sheet, address: "A:XFD" }] });
    expect(planUndo("find_replace", { sheet, address: "B2:B9", find: "a", replace: "b", apply: true })).toMatchObject({ targets: [{ sheet, address: "B2:B9" }] });
  });

  it("set_cells covers the box around its cells plus the copy target, and notices styles, borders, resizes and notes", () => {
    const plan = planUndo("set_cells", { sheet, cells: { B2: { value: 1 }, D5: { formula: "=B2", style: { bold: true } } }, copyTo: "B2:D9", resizeWidth: { type: "autofit" } });
    expect(plan).toMatchObject({
      kind: "cells", style: true, merged: true,
      targets: [{ sheet, address: "B2:D5" }, { sheet, address: "B2:D9" }],
      layout: { columns: [2, 4], rows: [2, 5], dimensions: ["width"] },
    });
    expect(planUndo("set_cells", { sheet, cells: { A1: { value: 1 } } })).toMatchObject({ style: false });
    expect(planUndo("set_cells", { sheet, cells: { A1: { borders: { top: { style: "solid", weight: "thin" } } } } })).toMatchObject({ style: true });
    expect(planUndo("set_cells", { sheet, cells: { A1: { note: "n" } } })).toMatchObject({ caveats: ["cell notes are not restored"] });
    expect(planUndo("set_cells", { sheet, cells: { A1: { value: 1 } }, copyTo: "A1:A5" })).toMatchObject({ caveats: [expect.stringContaining("copied along")] });
    expect(planUndo("set_cells", { sheet, cells: { nope: { value: 1 } } })).toMatchObject({ kind: "irreversible" });
  });

  it("works out which rows insert_rows adds", () => {
    expect(planUndo("insert_rows", { sheet, at: "5", count: 2, position: "above" })).toEqual({ kind: "insert_rows", sheet, row: 5, count: 2 });
    expect(planUndo("insert_rows", { sheet, at: "A5", count: 3, position: "below" })).toEqual({ kind: "insert_rows", sheet, row: 6, count: 3 });
    expect(planUndo("insert_rows", { sheet, at: "x", count: 1 })).toMatchObject({ kind: "irreversible" });
  });

  it("says plainly which workbook edits it cannot reverse, and every write tool has a plan", () => {
    for (const tool of ["execute_office_js", "modify_structure", "manage_sheet", "modify_chart", "modify_pivot_table", "modify_table", "add_shape", "delete_shapes", "conditional_format", "data_validation", "filter_range", "page_layout", "manage_names", "insert_sheets_from_file", "record_test_results"]) {
      expect(planUndo(tool, {})).toMatchObject({ kind: "irreversible", reason: expect.any(String) });
    }
    for (const tool of WRITE_TOOLS) {
      const plan = planUndo(tool, {});
      expect(["cells", "insert_rows", "none", "irreversible"]).toContain(plan.kind);
    }
  });
});

describe("photo helpers", () => {
  it("knows which text Excel would coerce when it is written back", () => {
    for (const text of ["00123", "12", "-5", "+1", "=A1", "'x", "1E3", "$5", "(3)", "#N/A", "TRUE", "false", "Jan 5", "Mar-20", ".5", "@home"]) expect(needsTextPrefix(text), text).toBe(true);
    for (const text of ["", "hello", "apple pie", "Name", "Xin chào"]) expect(needsTextPrefix(text), text).toBe(false);
  });

  it("describes a cell's look the same way regardless of Excel's bookkeeping and empty borders", () => {
    const base = { format: { font: { bold: true, name: "Calibri" }, fill: { color: "#FFFF00" }, horizontalAlignment: "Left", borders: { top: { style: "Continuous", color: "#000000", weight: "Thin" }, bottom: {} } } };
    const noisy = { "@odata.type": "x", format: { "@odata.type": "y", ...base.format, font: { "@odata.type": "z", ...base.format.font } } };
    expect(canonicalStyle(noisy)).toBe(canonicalStyle(base));
    expect(JSON.parse(canonicalStyle(base)).borders).toEqual({ top: { style: "Continuous", color: "#000000", weight: "Thin" } });
    expect(canonicalStyle({ format: { fill: { color: "" } } })).not.toBe(canonicalStyle(base));
  });

  it("tells a number from the same digits typed as text, and one format from another", () => {
    expect(cellHash(123, "General", undefined)).not.toBe(cellHash("123", "General", undefined));
    expect(cellHash("a", "General", undefined)).not.toBe(cellHash("a", "0.00", undefined));
    expect(cellHash("a", "General", "{}")).not.toBe(cellHash("a", "General", undefined));
    expect(cellHash("a", "General", undefined)).toBe(cellHash("a", "General", undefined));
  });

  it("fingerprints a photo cell by cell", () => {
    const a = photo("a");
    expect(sameFingerprint(fingerprintOf(a), fingerprintOf(photo("a")))).toBe(true);
    expect(sameFingerprint(fingerprintOf(a), fingerprintOf(photo("b")))).toBe(false);
  });
});

function photo(value: string, cells = 1) {
  return {
    blocks: [{ sheetId: "s1", sheetName: "Data", row: 0, col: 0, rows: 1, cols: cells, formulas: [Array(cells).fill(value)], valueTypes: [Array(cells).fill("String")], numberFormat: [Array(cells).fill("General")] }],
    layouts: [],
    tables: [],
  };
}

const ok: ToolResult = { ok: true, data: {} };
const refused: ToolResult = { ok: false, error: { code: "would_overwrite", message: "no" } };
const ctx = (callId: string, requestId = "r1") => ({ callId, requestId });

/** Wire the mocks so a tool "changes" the photographed cells (fingerprint differs) or not. */
function photographs(value = "before", changes = true) {
  mocked.photograph.mockResolvedValue(photo(value));
  mocked.fingerprintNow.mockResolvedValue(fingerprintOf(photo(changes ? "after" : value)));
}

describe("UndoJournal", () => {
  let backup: ReturnType<typeof vi.fn>;
  let journal: UndoJournal;
  beforeEach(() => {
    vi.resetAllMocks();
    backup = vi.fn(async () => ({ file: "Book1-1.xlsx", folder: "C:\\backups" }));
    journal = new UndoJournal((tool) => WRITE_TOOLS.has(tool), backup);
  });

  it("passes reads straight through without touching Excel", async () => {
    const exec = vi.fn(async () => ok);
    await journal.run(ctx("c1"), "read_range", { address: "A1" }, exec);
    expect(exec).toHaveBeenCalledTimes(1);
    expect(mocked.photograph).not.toHaveBeenCalled();
    expect(journal.entriesFor("c1")).toEqual([]);
  });

  it("photographs before and fingerprints after a cell edit, and keeps it undoable", async () => {
    photographs();
    const order: string[] = [];
    mocked.photograph.mockImplementation(async () => { order.push("photo"); return photo("before"); });
    const exec = vi.fn(async () => { order.push("tool"); return ok; });
    const result = await journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [["x"]] }, exec);
    expect(result).toBe(ok);
    expect(order).toEqual(["photo", "tool"]);
    expect(journal.stateFor("c1")).toEqual({ state: "undoable" });
    expect(journal.entriesFor("c1")[0]).toMatchObject({ tool: "write_range", requestId: "r1", cells: 1 });
  });

  it("leaves no entry for a refused call that changed nothing, but keeps one if it half-ran", async () => {
    photographs("v", false);
    await journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [["x"]] }, async () => refused);
    expect(journal.stateFor("c1")).toEqual({ state: "none" });
    photographs("v", true);
    await journal.run(ctx("c2"), "write_range", { sheet, address: "A1", values: [["x"]] }, async () => refused);
    expect(journal.stateFor("c2")).toEqual({ state: "undoable" });
  });

  it("does not stand in the way of tools that do not change anything", async () => {
    const exec = vi.fn(async () => ok);
    await journal.run(ctx("c1"), "find_replace", { find: "a", replace: "b" }, exec);
    expect(exec).toHaveBeenCalledTimes(1);
    expect(mocked.photograph).not.toHaveBeenCalled();
    expect(journal.stateFor("c1")).toEqual({ state: "none" });
  });

  it("flags edits it cannot reverse, taking one workbook copy per reply before the first of them", async () => {
    const exec = vi.fn(async () => ok);
    await journal.run(ctx("c1", "r1"), "execute_office_js", { code: "x" }, exec);
    await journal.run(ctx("c2", "r1"), "modify_chart", {}, exec);
    await journal.run(ctx("c3", "r2"), "manage_sheet", { operation: "delete" }, exec);
    expect(backup).toHaveBeenCalledTimes(2);
    expect(journal.stateFor("c1")).toMatchObject({ state: "unavailable", reason: expect.stringContaining("free-form"), backup: { file: "Book1-1.xlsx" } });
    expect(journal.stateFor("c2")).toMatchObject({ state: "unavailable", backup: { file: "Book1-1.xlsx" } });
  });

  it("takes the workbook copy before the irreversible tool runs, and still runs it when the copy fails", async () => {
    const order: string[] = [];
    backup.mockImplementationOnce(async () => { order.push("backup"); throw new Error("bridge offline"); });
    const result = await journal.run(ctx("c1"), "execute_office_js", { code: "x" }, async () => { order.push("tool"); return ok; });
    expect(order).toEqual(["backup", "tool"]);
    expect(result).toBe(ok);
    expect(journal.stateFor("c1")).toMatchObject({ state: "unavailable", backupError: "bridge offline" });
    expect(journal.stateFor("c1").backup).toBeUndefined();
  });

  it("gives calls made outside any reply their own workbook copy", async () => {
    const exec = vi.fn(async () => ok);
    await journal.run(ctx("c1", ""), "execute_office_js", { code: "x" }, exec);
    await journal.run(ctx("c2", "no-turn"), "execute_office_js", { code: "x" }, exec);
    await journal.run(ctx("c3", "no-turn"), "execute_office_js", { code: "x" }, exec);
    expect(backup).toHaveBeenCalledTimes(3);
    expect(isReply("no-turn")).toBe(false);
    expect(isReply("")).toBe(false);
    expect(isReply("abc")).toBe(true);
  });

  it("does not record a failed irreversible call", async () => {
    await journal.run(ctx("c1"), "execute_office_js", { code: "x" }, async () => refused);
    expect(journal.stateFor("c1")).toEqual({ state: "none" });
  });

  it("flags a change that is too big to photograph (with a copy), but still runs it", async () => {
    mocked.photograph.mockRejectedValue(new TooLarge("this change touches more than 20,000 cells"));
    const exec = vi.fn(async () => ok);
    const result = await journal.run(ctx("c1"), "format_range", { sheet, address: "A1:Z1000", format: { bold: true } }, exec);
    expect(result).toBe(ok);
    expect(journal.stateFor("c1")).toMatchObject({ state: "unavailable", reason: expect.stringContaining("20,000 cells"), backup: { file: "Book1-1.xlsx" } });
  });

  it("still runs the tool when photographing fails for any other reason", async () => {
    mocked.photograph.mockRejectedValue(new Error("InactiveWorkbook"));
    const exec = vi.fn(async () => ok);
    await journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [[1]] }, exec);
    expect(exec).toHaveBeenCalledTimes(1);
    expect(journal.stateFor("c1")).toMatchObject({ state: "unavailable", reason: expect.stringContaining("InactiveWorkbook") });
  });

  it("records inserted rows using where the tool says it inserted them", async () => {
    mocked.locateInsert.mockResolvedValue({ sheetId: "s1", sheetName: "Data", row: 5, count: 2 });
    await journal.run(ctx("c1"), "insert_rows", { sheet, at: "5", count: 2, position: "above" }, async () => ({ ok: true, data: { shiftedRange: "6:8" } }));
    expect(journal.entriesFor("c1")[0]!.inserted).toEqual({ sheetId: "s1", sheetName: "Data", row: 6, count: 3 });
    mocked.removeInsertedRows.mockResolvedValue(null);
    expect(await journal.undoCall("c1")).toEqual({ ok: true, message: "Removed 3 inserted rows." });
    expect(mocked.removeInsertedRows).toHaveBeenCalledWith({ sheetId: "s1", sheetName: "Data", row: 6, count: 3 });
  });

  it("refuses to remove inserted rows that now hold data and stays undoable", async () => {
    mocked.locateInsert.mockResolvedValue({ sheetId: "s1", sheetName: "Data", row: 5, count: 1 });
    await journal.run(ctx("c1"), "insert_rows", { sheet, at: "5", count: 1 }, async () => ({ ok: true, data: { shiftedRange: "5:5" } }));
    mocked.removeInsertedRows.mockResolvedValue('the rows inserted on "Data" now hold data');
    const outcome = await journal.undoCall("c1");
    expect(outcome.ok).toBe(false);
    expect(outcome.message).toContain("now hold data");
    expect(journal.stateFor("c1").state).toBe("undoable");
  });

  it("undoes a reply newest first and marks it undone, once", async () => {
    photographs();
    await journal.run(ctx("c1", "r1"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    await journal.run(ctx("c2", "r1"), "clear_range", { sheet, address: "A1" }, async () => ok);
    await journal.run(ctx("c3", "r2"), "write_range", { sheet, address: "A2", values: [[1]] }, async () => ok);
    const order: string[] = [];
    mocked.restorePhoto.mockImplementation(async () => { order.push(String(order.length)); return { restoredCells: 1, conflicts: [], conflictCount: 0, unrestored: [], unrestoredCount: 0, missingCells: 0, notes: [] }; });
    const outcome = await journal.undoRequest("r1");
    expect(outcome.ok).toBe(true);
    expect(mocked.restorePhoto).toHaveBeenCalledTimes(2);
    expect(journal.stateFor("c1").state).toBe("undone");
    expect(journal.stateFor("c2").state).toBe("undone");
    expect(journal.stateFor("c3").state).toBe("undoable");
    expect(await journal.undoRequest("r1")).toEqual({ ok: false, message: "There is nothing to undo here." });
    expect(mocked.restorePhoto).toHaveBeenCalledTimes(2);
  });

  it("restores entries in reverse order of the changes", async () => {
    photographs();
    mocked.photograph.mockImplementation(async () => photo("before")); // a new photo object per call
    await journal.run(ctx("c1", "r1"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    await journal.run(ctx("c2", "r1"), "write_range", { sheet, address: "A2", values: [[1]] }, async () => ok);
    const seen: number[] = [];
    mocked.restorePhoto.mockImplementation(async (p) => { seen.push(journal.entriesOf("r1").find((e) => e.photo === p)!.id); return { restoredCells: 1, conflicts: [], conflictCount: 0, unrestored: [], unrestoredCount: 0, missingCells: 0, notes: [] }; });
    await journal.undoRequest("r1");
    expect(seen).toEqual([2, 1]);
  });

  it("keeps an entry undoable when Excel refuses the undo, and reports why", async () => {
    photographs();
    await journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    mocked.restorePhoto.mockRejectedValueOnce(new Error("The workbook being called by this API has lost focus."));
    const outcome = await journal.undoCall("c1");
    expect(outcome.ok).toBe(false);
    expect(outcome.message).toContain("lost focus");
    expect(journal.stateFor("c1").state).toBe("undoable");
  });

  it("reports a partly restored undo as not fully ok", async () => {
    photographs();
    await journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    mocked.restorePhoto.mockResolvedValue({ restoredCells: 3, conflicts: [], conflictCount: 0, unrestored: [{ sheet: "Data", address: "A2" }], unrestoredCount: 1, missingCells: 0, notes: [] });
    const outcome = await journal.undoCall("c1");
    expect(outcome.ok).toBe(false);
    expect(outcome.message).toContain("Data!A2");
  });

  it("shows what a plan could not put back after undoing", async () => {
    photographs();
    await journal.run(ctx("c1"), "clear_range", { sheet, address: "A1", clearType: "all" }, async () => ok);
    mocked.restorePhoto.mockResolvedValue({ restoredCells: 1, conflicts: [], conflictCount: 0, unrestored: [], unrestoredCount: 0, missingCells: 0, notes: [] });
    expect((await journal.undoCall("c1")).message).toContain("Note: notes, hyperlinks");
  });

  it("runs one tool at a time, so a photo is never taken in the middle of another edit", async () => {
    photographs();
    const events: string[] = [];
    mocked.photograph.mockImplementation(async () => { events.push("photo"); return photo("before"); });
    const slow = (name: string) => async () => { events.push(`${name}:start`); await new Promise((r) => setTimeout(r, 15)); events.push(`${name}:end`); return ok; };
    await Promise.all([
      journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [[1]] }, slow("one")),
      journal.run(ctx("c2"), "write_range", { sheet, address: "A2", values: [[1]] }, slow("two")),
    ]);
    expect(events).toEqual(["photo", "one:start", "one:end", "photo", "two:start", "two:end"]);
  });

  it("keeps working after a tool throws", async () => {
    photographs();
    await expect(journal.run(ctx("c1"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => { throw new Error("boom"); })).rejects.toThrow("boom");
    const result = await journal.run(ctx("c2"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    expect(result).toBe(ok);
    expect(journal.stateFor("c2").state).toBe("undoable");
  });

  it("forgets the oldest photos once the history is full, and says why", async () => {
    mocked.photograph.mockResolvedValue(photo("before", JOURNAL_LIMITS.cells / 2 + 1));
    mocked.fingerprintNow.mockResolvedValue(fingerprintOf(photo("after", JOURNAL_LIMITS.cells / 2 + 1)));
    await journal.run(ctx("old"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    await journal.run(ctx("new"), "write_range", { sheet, address: "A1", values: [[1]] }, async () => ok);
    expect(journal.stateFor("old")).toMatchObject({ state: "unavailable", reason: expect.stringContaining("too old") });
    expect(journal.stateFor("new").state).toBe("undoable");
    expect(journal.entriesFor("old")[0]!.photo).toBeUndefined();
  });

  it("describes a restore in plain words", () => {
    const base = { restoredCells: 1, conflicts: [], conflictCount: 0, unrestored: [], unrestoredCount: 0, missingCells: 0, notes: [] };
    expect(describeRestore(base)).toBe("Restored 1 cell.");
    expect(describeRestore({ ...base, restoredCells: 1234, conflictCount: 2, conflicts: [{ sheet: "Data", address: "B2" }, { sheet: "Data", address: "B3" }] }))
      .toBe("Restored 1,234 cells. 2 cells were changed after the assistant's edit and left as they are (Data!B2, Data!B3).");
    expect(describeRestore({ ...base, missingCells: 4 })).toContain("4 cells could not be restored because their sheet was deleted");
    expect(describeRestore({ ...base, notes: ["merged cells were changed"] })).toContain("Note: merged cells were changed.");
  });
});
