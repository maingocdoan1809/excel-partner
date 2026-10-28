import { describe, expect, it } from "vitest";
import { SelectionSnapshotSchema } from "../protocol/src/index.js";
import { mergeWindowsSelection, needsNativeRead } from "../bridge/src/windows-selection.js";
import { buildTurnPrompt } from "../providers/src/prompt.js";

const cells = SelectionSnapshotSchema.parse({ workbookName: "Flow.xlsx", sheet: "Flow", sheetId: "sheet-1",
  address: "B2", rowCount: 1, columnCount: 1, values: [["stale cell"]], headers: ["Header"], truncated: false });
const shapes = [{ name: "Decision", nativeId: "2", nativeType: 1, connector: false },
  { name: "Arrow", nativeId: "3", nativeType: 9, connector: true }];

describe("Windows multi-selection", () => {
  it("passes the full selection through protocol validation and into the model prompt without stale cells", () => {
    const selection = SelectionSnapshotSchema.parse(mergeWindowsSelection(cells, {
      status: "ok", workbookName: "Flow.xlsx", sheet: "Flow", shapes,
    }));
    expect(selection).toMatchObject({ selectedShapes: shapes, selectedShapesStatus: "complete", selectedShapesSource: "windows_com",
      address: "", values: [], headers: null, rowCount: 0, columnCount: 0 });
    const context = JSON.parse(buildTurnPrompt("check these", selection).split("\n")[1]!);
    expect(context.selectedShapes).toEqual(shapes);
    expect(context.selectedShapesStatus).toBe("complete");
    expect(context.values).toEqual([]);
  });

  it("rejects identities from another workbook or worksheet", () => {
    for (const identity of [{ workbookName: "Other.xlsx", sheet: "Flow" }, { workbookName: "Flow.xlsx", sheet: "Other" }]) {
      const result = mergeWindowsSelection(cells, { status: "ok", ...identity, shapes });
      expect(result.selectedShapes).toBeUndefined();
      expect(result.selectedShapesStatus).toBe("unavailable");
      expect(result.values).toEqual(cells.values);
    }
  });

  it("removes outdated group identities when native reading fails", () => {
    const before = mergeWindowsSelection(cells, { status: "ok", workbookName: "Flow.xlsx", sheet: "Flow", shapes });
    const after = mergeWindowsSelection(before, { status: "unavailable", reason: "Excel is busy" });
    expect(after.selectedShapes).toBeUndefined();
    expect(after.selectedShapesStatus).toBe("unavailable");
    expect(buildTurnPrompt("check", after)).toContain("Excel is busy");
  });

  it("keeps an active shape only when it belongs to the native selection", () => {
    const activeShape = { name: "Decision", type: "GeometricShape", text: "Yes?", left: 0, top: 0, width: 10, height: 10 };
    const reply = { status: "ok" as const, workbookName: "Flow.xlsx", sheet: "Flow", shapes };
    expect(mergeWindowsSelection({ ...cells, activeShape }, reply).activeShape).toEqual(activeShape);
    expect(mergeWindowsSelection({ ...cells, activeShape: { ...activeShape, name: "Old" } }, reply).activeShape).toBeUndefined();
    expect(mergeWindowsSelection({ ...cells, activeShape }, { status: "not_shapes" }).selectedShapesStatus).toBe("unavailable");
  });

  it("retains regular cell data when COM confirms there are no shapes", () => {
    expect(mergeWindowsSelection(cells, { status: "not_shapes" })).toEqual(cells);
  });
});

describe("when the Windows COM read runs", () => {
  // Reading a selected chart through COM (Selection.ShapeRange) crashed Excel, and the chart is already identified by Office.js.
  const active = (type: string) => SelectionSnapshotSchema.parse({ ...cells, address: "", rowCount: 0, columnCount: 0, values: [], headers: null,
    activeShape: { name: "MsgPerDay", type, text: null, left: 15, top: 230, width: 354, height: 240 }, shapeSelectionStatus: "active" });

  it("skips plain cell selections and selected charts", () => {
    expect(needsNativeRead(cells)).toBe(false);
    expect(needsNativeRead(active("Chart"))).toBe(false);
  });

  it("still reads selected shapes and connectors", () => {
    expect(needsNativeRead(active("Rectangle"))).toBe(true);
    expect(needsNativeRead(active("Line"))).toBe(true);
  });
});
