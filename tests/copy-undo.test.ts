import { beforeEach, expect, it } from "vitest";
import { executeTool, WRITE_TOOLS } from "../addin/taskpane/executor/index.js";
import { UndoJournal } from "../addin/taskpane/executor/journal.js";
import { photograph } from "../addin/taskpane/executor/photo.js";
import { planUndo } from "../addin/taskpane/executor/undo-plan.js";
import { parseToolInput } from "../protocol/src/index.js";
import { MockWorkbook, installOfficeMock } from "./support/office-mock.js";

let book: MockWorkbook;
beforeEach(() => {
  book = new MockWorkbook(); book.addSheet("Sheet1");
  book.selection = { sheet: "Sheet1", address: "A1" }; installOfficeMock(book);
});

it("undo restores cells beyond an anchor destination using the actual photograph and restore code", async () => {
  book.fill("Sheet1", [["one"], ["two"]], "A1");
  book.fill("Sheet1", [["keep"]], "D2");
  const journal = new UndoJournal(tool => WRITE_TOOLS.has(tool));
  const args = parseToolInput("copy_to", { source: "A1:A2", destination: "D1", copyType: "values", allowOverwrite: true });
  expect((await journal.run({ callId: "copy", requestId: "r" }, "copy_to", args, executeTool)).ok).toBe(true);
  expect(book.get("Sheet1", "D2")).toBe("two");
  expect(journal.stateFor("copy").state).toBe("undoable");
  expect((await journal.undoCall("copy")).ok).toBe(true);
  expect(book.get("Sheet1", "D1")).toBeNull();
  expect(book.get("Sheet1", "D2")).toBe("keep");
});

it("set_cells photographs its expanded copy destination before any cells change", async () => {
  book.fill("Sheet1", [["keep"]], "E2");
  const plan = planUndo("set_cells", { cells: { A1: { value: "one" }, B2: { value: "two" } }, copyTo: "D1" });
  if (plan.kind !== "cells") throw new Error("Expected a cell plan");
  // Inspect cell geometry; style/merge restoration has independent coverage.
  const photo = await photograph({ ...plan, style: false, merged: false });
  expect(photo.blocks).toContainEqual(expect.objectContaining({ row: 0, col: 3, rows: 2, cols: 2 }));
});
