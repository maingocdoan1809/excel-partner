import { beforeEach, describe, expect, it } from "vitest";
import { executeTool } from "../addin/taskpane/executor/index.js";
import { buildFilterCriteria, filterCriterion, relativeColumn } from "../addin/taskpane/executor/data.js";
import { buildCellValueRule, buildValidationRule, ruleFormula } from "../addin/taskpane/executor/rules.js";
import { describeBorder, groupByRectangles } from "../addin/taskpane/executor/styles.js";
import { parseToolInput } from "../protocol/src/index.js";
import { MockSheet, MockWorkbook, installOfficeMock } from "./support/office-mock.js";

let book: MockWorkbook;
let sheet: MockSheet;

beforeEach(() => {
  book = new MockWorkbook();
  sheet = book.addSheet("Sheet1");
  book.fill("Sheet1", [
    ["Item", "Qty", "Price"],
    ["Pear", 2, 2],
    ["Apple", 3, 1.5],
    ["Fig", 1, 4],
  ]);
  book.selection = { sheet: "Sheet1", address: "A1" };
  installOfficeMock(book);
});

async function call(tool: string, args: unknown) {
  return executeTool(tool, parseToolInput(tool as never, args));
}
const data = (r: Awaited<ReturnType<typeof call>>) => (r.ok ? (r.data as any) : (() => { throw new Error(JSON.stringify(r.error)); })());
function expectError(r: Awaited<ReturnType<typeof call>>, code: string) {
  expect(r.ok).toBe(false);
  if (!r.ok) expect(r.error.code).toBe(code);
  return r.ok ? null : r.error;
}

describe("read_range paging", () => {
  it("pages rows with maxRows / offset", async () => {
    const first = data(await call("read_range", { address: "A1:C4", maxRows: 2 }));
    expect(first).toMatchObject({ returnedRows: 2, hasMore: true, nextOffset: 2, values: [["Item", "Qty", "Price"], ["Pear", 2, 2]] });
    const second = data(await call("read_range", { address: "A1:C4", maxRows: 2, offset: 2 }));
    expect(second).toMatchObject({ returnedRows: 2, hasMore: false, nextOffset: null, values: [["Apple", 3, 1.5], ["Fig", 1, 4]] });
  });

  it("returns csv pages too", async () => {
    const page = data(await call("read_range", { address: "A1:C4", format: "csv", maxRows: 2, offset: 1 }));
    expect(page.csv).toBe("Pear,2,2\nApple,3,1.5");
  });
});

describe("set_cells", () => {
  it("writes values, formulas, styles, borders and notes in one call and reports formula results", async () => {
    const result = await call("set_cells", {
      cells: {
        E1: { value: "Total", style: { bold: true, backgroundColor: "#DDEBF7", numberFormat: "@" } },
        E2: { formula: "=B2*C2", borders: { bottom: { weight: "medium", color: "#112233" } } },
        E3: { formula: "B3*C3", note: "computed" },
      },
    });
    expect(result).toMatchObject({ ok: true, data: { cellsWritten: 3, formulaResults: { E2: 4, E3: 4.5 } } });
    expect(book.get("Sheet1", "E1")).toBe("Total");
    expect(book.get("Sheet1", "E3")).toBe("=B3*C3");
    expect(sheet.noteTexts.get("E3")).toBe("computed");
    expect(sheet.styles.get("0,4")).toMatchObject({ "font.bold": true, "fill.color": "#DDEBF7", numberFormat: "@" });
    expect(sheet.styles.get("1,4")).toMatchObject({ "border.EdgeBottom.weight": "Medium", "border.EdgeBottom.color": "#112233", "border.EdgeBottom.style": "Continuous" });
  });

  it("refuses to overwrite and changes nothing", async () => {
    const error = expectError(await call("set_cells", { cells: { E1: { value: "ok" }, A2: { value: "Mango" } } }), "would_overwrite");
    expect(error?.message).toMatch(/1 non-empty cell: A2/);
    expect(book.get("Sheet1", "A2")).toBe("Pear");
    expect(book.get("Sheet1", "E1")).toBeNull();
  });

  it("allows style-only changes on occupied cells without allowOverwrite", async () => {
    const result = await call("set_cells", { cells: { A1: { style: { bold: true } } } });
    expect(result.ok).toBe(true);
    expect(book.get("Sheet1", "A1")).toBe("Item");
  });

  it("copyTo fills a pattern down and adjusts relative references", async () => {
    const result = await call("set_cells", { cells: { D2: { formula: "=B2*C2" } }, copyTo: "D2:D4" });
    expect(result.ok).toBe(true);
    expect(book.get("Sheet1", "D4")).toBe("=B4*C4");
    expect(data(result).formulaResults).toMatchObject({ D2: 4, D3: 4.5, D4: 4 });
  });

  it("treats copyTo over existing data as an overwrite", async () => {
    book.fill("Sheet1", [["keep"]], "D3");
    expectError(await call("set_cells", { cells: { D2: { formula: "=B2" } }, copyTo: "D2:D4" }), "would_overwrite");
    expect(book.get("Sheet1", "D3")).toBe("keep");
  });

  it("surfaces formula errors in messages", async () => {
    const result = await call("set_cells", { cells: { E5: { formula: "=B2/0" } } });
    expect(result.ok).toBe(true);
  });

  it("rejects value + formula on one cell and bad addresses at validation", () => {
    expect(() => parseToolInput("set_cells", { cells: { A1: { value: 1, formula: "=1" } } })).toThrow();
    expect(() => parseToolInput("set_cells", { cells: { "A1:B2": { value: 1 } } })).toThrow();
    expect(() => parseToolInput("set_cells", { cells: {} })).toThrow();
  });

  it("fails before writing when notes are unsupported", async () => {
    book.supportsNotes = false;
    expectError(await call("set_cells", { cells: { E1: { value: 1, note: "n" } } }), "unsupported");
    expect(book.get("Sheet1", "E1")).toBeNull();
  });
});

describe("get_cell_ranges", () => {
  it("returns values, formulas and notes keyed by address, skipping empty cells", async () => {
    await call("set_cells", { cells: { D2: { formula: "=B2*C2", note: "calc" } } });
    const result = data(await call("get_cell_ranges", { ranges: ["A1:D3"], includeStyles: false }));
    expect(result.dimension).toBe("A1:D4");
    expect(result.cells).toMatchObject({ A1: "Item", B2: 2, D2: [4, "=B2*C2", "calc"] });
    expect(result.cells.D1).toBeUndefined();
    expect(result.hasMore).toBe(false);
  });

  it("groups identical styles into rectangles and reports borders", async () => {
    await call("format_range", { address: "A1:C1", format: { bold: true, fillColor: "#DDEBF7" } });
    await call("set_cells", { cells: { A2: { borders: { all: { weight: "thin" } } }, B2: { borders: { all: { weight: "thin" } } } } , allowOverwrite: true });
    const result = data(await call("get_cell_ranges", { ranges: ["A1:C3"] }));
    expect(result.styles["A1:C1"]).toMatchObject({ bold: true, backgroundColor: "#ddebf7" });
    expect(result.borders["A2:B2"]).toMatchObject({ top: "thin solid #000000", bottom: "thin solid #000000" });
    expect(result.baseline).toMatchObject({ fontFamily: "Calibri", fontSize: 11 });
  });

  it("pages through nextRanges when cellLimit is reached", async () => {
    const first = data(await call("get_cell_ranges", { ranges: ["A1:C4", "A6:A7"], includeStyles: false, cellLimit: 5 }));
    expect(first.hasMore).toBe(true);
    expect(Object.keys(first.cells).length).toBeGreaterThanOrEqual(5);
    expect(first.nextRanges).toEqual(["A3:C4", "A6:A7"]);
    const rest = data(await call("get_cell_ranges", { ranges: first.nextRanges, includeStyles: false }));
    expect(rest.cells.A3).toBe("Apple");
    expect(rest.hasMore).toBe(false);
  });

  it("clips whole-column references to the used range", async () => {
    const result = data(await call("get_cell_ranges", { ranges: ["B:B"], includeStyles: false }));
    expect(Object.keys(result.cells)).toEqual(["B1", "B2", "B3", "B4"]);
  });
});

describe("copy_to / resize_range / read_range_image / merge_range / sort_range", () => {
  it("copies with relative references and refuses to overwrite", async () => {
    await call("set_formula", { address: "D2", formula: "B2*C2" });
    const copied = await call("copy_to", { source: "D2", destination: "D3:D4" });
    expect(copied.ok).toBe(true);
    expect(book.get("Sheet1", "D4")).toBe("=B4*C4");
    expectError(await call("copy_to", { source: "A2:C2", destination: "A3:C4" }), "would_overwrite");
    expect((await call("copy_to", { source: "A1:C1", destination: "A3:C3", copyType: "formats" })).ok).toBe(true);
  });

  it("resizes columns and rows by points, standard and autofit", async () => {
    await call("resize_range", { address: "A:C", width: { type: "points", value: 80 }, height: { type: "autofit" } });
    const log = sheet.formatLog.flatMap((e) => Object.entries(e.props)).map(([k, v]) => `${k}=${v}`);
    expect(log).toContain("columnWidth=80");
    expect(log).toContain("autofitRows=true");
    await call("resize_range", { width: { type: "standard" } });
    expect(sheet.formatLog.some((e) => e.props["useStandardWidth"] === true)).toBe(true);
    expect(() => parseToolInput("resize_range", { address: "A:A" })).toThrow();
    expect(() => parseToolInput("resize_range", { width: { type: "points" } })).toThrow();
  });

  it("returns a rendered image payload", async () => {
    const result = data(await call("read_range_image", { address: "A1:C4" }));
    expect(result.image).toEqual({ mimeType: "image/png", data: "BASE64PNG" });
  });

  it("merges only when nothing but the top-left cell holds data", async () => {
    expectError(await call("merge_range", { address: "A1:B1", operation: "merge" }), "would_overwrite");
    expect((await call("merge_range", { address: "E1:F1", operation: "merge" })).ok).toBe(true);
    expect((await call("merge_range", { address: "A1:B1", operation: "merge", allowOverwrite: true })).ok).toBe(true);
    expect(book.get("Sheet1", "B1")).toBeNull();
    expectError(await call("merge_range", { address: "H1", operation: "merge" }), "invalid_arguments");
  });

  it("sorts by sheet column letters mapped into the range and rejects columns outside it", async () => {
    const result = await call("sort_range", { address: "A1:C4", keys: [{ column: "B", ascending: false }] });
    expect(result.ok).toBe(true);
    expect(sheet.sortLog[0]).toMatchObject({ fields: [{ key: 1, ascending: false }], hasHeaders: true });
    expect(book.get("Sheet1", "A2")).toBe("Apple"); // qty 3 first
    expectError(await call("sort_range", { address: "B1:C4", keys: [{ column: "A" }] }), "invalid_arguments");
  });
});

describe("manage_sheet", () => {
  it("adds, renames, moves, colours and hides sheets", async () => {
    expect(data(await call("manage_sheet", { operation: "add", name: "Data" }))).toMatchObject({ sheet: "Data" });
    expectError(await call("manage_sheet", { operation: "add", name: "Data" }), "invalid_arguments");
    await call("manage_sheet", { operation: "rename", sheet: { name: "Data" }, newName: "Raw" });
    expect(book.sheets.map((s) => s.name)).toEqual(["Sheet1", "Raw"]);
    await call("manage_sheet", { operation: "set_tab_color", sheet: { name: "Raw" }, tabColor: "#FF0000" });
    expect(book.sheet("Raw").tabColor).toBe("#FF0000");
    await call("manage_sheet", { operation: "hide", sheet: { name: "Raw" } });
    expect(book.sheet("Raw").visibility).toBe("Hidden");
    expectError(await call("manage_sheet", { operation: "hide", sheet: { name: "Sheet1" } }), "invalid_arguments");
  });

  it("refuses to delete without confirmation and says what would be lost", async () => {
    book.addSheet("Scratch");
    book.fill("Scratch", [["x", "y"]]);
    const error = expectError(await call("manage_sheet", { operation: "delete", sheet: { name: "Scratch" } }), "confirmation_required");
    expect(error?.message).toMatch(/A1:B1/);
    expect(book.sheets).toHaveLength(2);
    expect((await call("manage_sheet", { operation: "delete", sheet: { name: "Scratch" }, confirmDelete: true })).ok).toBe(true);
    expect(book.sheets.map((s) => s.name)).toEqual(["Sheet1"]);
    expectError(await call("manage_sheet", { operation: "delete", sheet: { name: "Sheet1" }, confirmDelete: true }), "invalid_arguments");
  });

  it("validates names and required fields", () => {
    expect(() => parseToolInput("manage_sheet", { operation: "rename" })).toThrow();
    expect(() => parseToolInput("manage_sheet", { operation: "add", name: "a/b" })).toThrow();
    expect(() => parseToolInput("manage_sheet", { operation: "delete" })).toThrow();
    expect(() => parseToolInput("manage_sheet", { operation: "move" })).toThrow();
  });

  it("duplicates a sheet", async () => {
    const result = data(await call("manage_sheet", { operation: "duplicate", sheet: { name: "Sheet1" }, name: "Copy" }));
    expect(result.duplicate).toBe("Copy");
    expect(book.get("Copy", "A2")).toBe("Pear");
  });
});

describe("structure: grouping", () => {
  it("groups, ungroups and collapses to an outline level", async () => {
    await call("modify_structure", { operation: "group", dimension: "rows", reference: "2", count: 3 });
    await call("modify_structure", { operation: "ungroup", dimension: "columns", reference: "B", count: 2 });
    await call("modify_structure", { operation: "show_levels", dimension: "rows", count: 1 });
    expect(sheet.groupLog).toEqual([
      { op: "group", option: "ByRows", address: "A2:XFD4" },
      { op: "ungroup", option: "ByColumns", address: "B1:C1048576" },
    ]);
    expect(sheet.outlineLog).toEqual([[1, 0]]);
    expect(() => parseToolInput("modify_structure", { operation: "group", dimension: "rows" })).toThrow();
    expect(() => parseToolInput("modify_structure", { operation: "show_levels", dimension: "rows", count: 9 })).toThrow();
  });
});

describe("filters", () => {
  it("applies value, condition and top filters on the right relative column", async () => {
    await call("filter_range", { operation: "apply", address: "A1:C4", column: "C", criteria: { values: ["2", "4"] } });
    await call("filter_range", { operation: "apply", address: "A1:C4", column: "B", criteria: { conditions: [{ operator: "greaterThan", value: 1 }, { operator: "lessThan", value: 3 }], join: "and" } });
    const log = sheet.autoFilterState.log as any[];
    expect(log[0]).toMatchObject({ column: 2, criteria: { filterOn: "Values", values: ["2", "4"] } });
    expect(log[1]).toMatchObject({ column: 1, criteria: { filterOn: "Custom", criterion1: ">1", criterion2: "<3", operator: "And" } });
    await call("filter_range", { operation: "clear" });
    await call("filter_range", { operation: "remove" });
    expect(sheet.autoFilterState.enabled).toBe(false);
  });

  it("removes the old filter when the range changes", async () => {
    await call("filter_range", { operation: "apply", address: "A1:C4" });
    await call("filter_range", { operation: "apply", address: "A1:B4" });
    expect((sheet.autoFilterState.log as any[]).map((l) => l.op)).toEqual(["apply", "remove", "apply"]);
  });

  it("validates apply arguments", () => {
    expect(() => parseToolInput("filter_range", { operation: "apply" })).toThrow();
    expect(() => parseToolInput("filter_range", { operation: "apply", address: "A1:C4", column: "B" })).toThrow();
  });
});

describe("pure builders", () => {
  it("maps filter conditions to Excel criteria strings", () => {
    expect(filterCriterion("contains", "ab")).toBe("=*ab*");
    expect(filterCriterion("notContains", "ab")).toBe("<>*ab*");
    expect(filterCriterion("beginsWith", "ab")).toBe("=ab*");
    expect(filterCriterion("endsWith", "ab")).toBe("=*ab");
    expect(filterCriterion("blank")).toBe("=");
    expect(filterCriterion("nonBlank")).toBe("<>");
    expect(buildFilterCriteria({ top: { kind: "topItems", n: 5 } })).toMatchObject({ filterOn: "TopItems", criterion1: "5" });
    expect(() => buildFilterCriteria({ conditions: [{ operator: "equals" }], join: "and" })).toThrow(/needs a value/);
  });

  it("maps relative columns and rejects outside ones", () => {
    expect(relativeColumn("C", 1, 3, "B1:D9")).toBe(1);
    expect(() => relativeColumn("A", 1, 3, "B1:D9")).toThrow(/outside/);
  });

  it("builds conditional-format rule bounds", () => {
    expect(ruleFormula(5)).toBe("=5");
    expect(ruleFormula("10")).toBe("=10");
    expect(ruleFormula("=$A$1")).toBe("=$A$1");
    expect(ruleFormula('say "hi"')).toBe('="say ""hi"""');
    expect(buildCellValueRule({ type: "cellValue", operator: "greaterThanOrEqual", value: 100, format: { bold: true } })).toEqual({
      formula1: "=100",
      operator: "GreaterThanOrEqual",
    });
  });

  it("builds validation rules for each kind", () => {
    expect(buildValidationRule({ type: "list", source: ["a", "b"], inCellDropDown: true })).toEqual({ list: { inCellDropDown: true, source: "a,b" } });
    expect(buildValidationRule({ type: "list", source: "=$A$1:$A$3", inCellDropDown: false })).toEqual({ list: { inCellDropDown: false, source: "=$A$1:$A$3" } });
    expect(buildValidationRule({ type: "whole", operator: "between", value1: 1, value2: 10 })).toEqual({ wholeNumber: { formula1: 1, formula2: 10, operator: "Between" } });
    expect(buildValidationRule({ type: "textLength", operator: "lessThan", value1: 5 })).toEqual({ textLength: { formula1: 5, operator: "LessThan" } });
    expect(buildValidationRule({ type: "custom", formula: "=ISNUMBER(A1)" })).toEqual({ custom: { formula: "=ISNUMBER(A1)" } });
    expect(() => buildValidationRule({ type: "list", source: ["a,b"], inCellDropDown: true })).toThrow(/commas/);
    expect(() => parseToolInput("data_validation", { address: "A1", operation: "set", rule: { type: "whole", operator: "between", value1: 1 } })).toThrow(/value2/);
  });

  it("groups cell styles into rectangles", () => {
    const bold = { bold: true as const };
    const grid = [
      [bold, bold, null],
      [bold, bold, { italic: true as const }],
      [null, null, { italic: true as const }],
    ];
    expect(groupByRectangles(grid, 0, 0)).toEqual({ "A1:B2": bold, "C2:C3": { italic: true } });
  });

  it("describes borders compactly", () => {
    expect(describeBorder({ style: "Continuous", weight: "Medium", color: "#112233" })).toBe("medium solid #112233");
    expect(describeBorder({ style: "None" })).toBeNull();
    expect(describeBorder({ style: "Dash", weight: "Hairline", color: "#FFFFFF" })).toBe("thin dashed #ffffff");
  });
});

describe("opt-in tool gating", () => {
  it("keeps execute_office_js out of the default tool list", async () => {
    const { exposedToolNames } = await import("../protocol/src/index.js");
    expect(exposedToolNames(false)).not.toContain("execute_office_js");
    expect(exposedToolNames(true)).toContain("execute_office_js");
  });
});
