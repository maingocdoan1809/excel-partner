import { beforeEach, describe, expect, it } from "vitest";
import { getSelectionLabel } from "../addin/taskpane/executor/read.js";
import { executeTool } from "../addin/taskpane/executor/index.js";
import { parseToolInput, SelectionSnapshotSchema } from "../protocol/src/index.js";
import { MockWorkbook, MockSheet, MockShape, installOfficeMock } from "./support/office-mock.js";

let book: MockWorkbook;
let sheet: MockSheet;

beforeEach(() => {
  book = new MockWorkbook();
  sheet = book.addSheet("Sheet1");
  book.fill("Sheet1", [
    ["Item", "Qty", "Price"],
    ["Apple", 3, 1.5],
    ["Pear", 2, 2],
  ]);
  book.selection = { sheet: "Sheet1", address: "A2:C2" };
  installOfficeMock(book);
});

/** Run a tool the way the bridge does: validate/default with the shared schema, then execute. */
async function call(tool: string, args: unknown) {
  return executeTool(tool, parseToolInput(tool as never, args));
}

function expectError(result: Awaited<ReturnType<typeof call>>, code: string) {
  expect(result.ok).toBe(false);
  if (!result.ok) expect(result.error.code).toBe(code);
  return result.ok ? null : result.error;
}

describe("reading", () => {
  it("reports the selection with a sheet-free address and detected headers", async () => {
    const result = await call("get_selection", {});
    expect(result).toMatchObject({
      ok: true,
      data: { sheet: "Sheet1", address: "A2:C2", values: [["Apple", 3, 1.5]], headers: ["Item", "Qty", "Price"] },
    });
  });

  it("reports a selected shape instead of failing", async () => {
    book.invalidSelection = true;
    const result = await call("get_selection", {});
    expect(result).toMatchObject({
      ok: true,
      data: { sheet: "Sheet1", address: "", rowCount: 0, values: [], note: expect.stringMatching(/shape, chart or other object/) },
    });
  });

  it("still fails for other Office errors", async () => {
    book.sheets.length = 0;
    book.active = 0;
    expect((await call("get_selection", {})).ok).toBe(false);
  });

  it("has no headers when the selection is the first row", async () => {
    book.selection = { sheet: "Sheet1", address: "A1:C1" };
    const result = await call("get_selection", {});
    expect(result.ok && (result.data as any).headers).toBeNull();
  });

  it("returns csv with proper quoting", async () => {
    book.fill("Sheet1", [['say "hi", ok', null, "line\nbreak"]], "A5");
    const result = await call("read_range", { address: "A5:C5", format: "csv" });
    expect(result.ok && (result.data as any).csv).toBe('"say ""hi"", ok",,"line\nbreak"');
  });

  it("reads formulas alongside values", async () => {
    await call("set_formula", { address: "D2", formula: "B2*C2" });
    const result = await call("read_range", { address: "D2", includeFormulas: true });
    expect(result.ok && (result.data as any)).toMatchObject({ values: [[4.5]], formulas: [["=B2*C2"]] });
  });

  it("reports a missing sheet with a typed error", async () => {
    expectError(await call("read_range", { address: "A1", sheet: { name: "Nope" } }), "sheet_not_found");
  });
});

describe("overwrite protection", () => {
  it("writes into empty cells without a flag", async () => {
    const result = await call("write_range", { address: "E1", values: [["a", "b"], ["c", "d"]] });
    expect(result).toMatchObject({ ok: true, data: { address: "E1:F2", rowsWritten: 2, columnsWritten: 2 } });
    expect(book.get("Sheet1", "F2")).toBe("d");
  });

  it("refuses to overwrite and leaves the workbook untouched", async () => {
    const error = expectError(await call("write_range", { address: "A2", values: [["X", "Y"], ["Z", "W"]] }), "would_overwrite");
    expect(error?.message).toMatch(/Would overwrite 4 non-empty cells: A2, B2, A3, B3/);
    expect(error?.message).toMatch(/allowOverwrite=true/);
    expect(book.get("Sheet1", "A2")).toBe("Apple");
    expect(book.get("Sheet1", "B3")).toBe(2);
  });

  it("allows the write once allowOverwrite is set", async () => {
    const result = await call("write_range", { address: "A2", values: [["Mango"]], allowOverwrite: true });
    expect(result.ok).toBe(true);
    expect(book.get("Sheet1", "A2")).toBe("Mango");
  });

  it("counts a formula as non-empty and ignores formatting-only cells", async () => {
    await call("set_formula", { address: "E1", formula: "1+1" });
    expectError(await call("write_range", { address: "E1", values: [[5]] }), "would_overwrite");

    await call("format_range", { address: "H1", format: { bold: true } });
    expect((await call("write_range", { address: "H1", values: [["ok"]] })).ok).toBe(true);
  });

  it("fails closed when the cells cannot be inspected", async () => {
    book.failFormulaReads = true;
    const result = await call("write_range", { address: "A2", values: [["Mango"]] });
    expect(result.ok).toBe(false);
    expect(book.get("Sheet1", "A2")).toBe("Apple");
  });

  it("applies the same guard to set_formula and reports how many cells collide", async () => {
    const error = expectError(await call("set_formula", { address: "B2:B3", formula: "1" }), "would_overwrite");
    expect(error?.message).toMatch(/2 non-empty cells: B2, B3/);
  });

  it("does not leave the first rows written when a later cell collides", async () => {
    book.fill("Sheet1", [["keep"]], "A6");
    expectError(await call("write_range", { address: "A4", values: [["n1"], ["n2"], ["n3"]] }), "would_overwrite");
    expect(book.get("Sheet1", "A4")).toBeNull();
    expect(book.get("Sheet1", "A5")).toBeNull();
  });
});

describe("write_range shape handling", () => {
  it("expands a single anchor cell to the size of the values", async () => {
    const result = await call("write_range", { address: "A10", values: [[1, 2, 3]] });
    expect(result.ok && (result.data as any).address).toBe("A10:C10");
  });

  it("rejects an address whose size differs from the values", async () => {
    const error = expectError(await call("write_range", { address: "A10:B11", values: [[1, 2, 3]] }), "invalid_arguments");
    expect(error?.message).toMatch(/2x2 but values are 1x3/);
    expect(book.get("Sheet1", "A10")).toBeNull();
  });

  it("pads ragged rows without erasing neighbours", async () => {
    book.fill("Sheet1", [["old"]], "C10");
    const result = await call("write_range", { address: "A10", values: [[1, 2, 3], [4]], mode: "top_left", allowOverwrite: true });
    expect(result.ok).toBe(true);
    expect(book.get("Sheet1", "A11")).toBe(4);
    expect(book.get("Sheet1", "B11")).toBeNull();
  });

  it("treats strings starting with = as formulas", async () => {
    await call("write_range", { address: "E1", values: [["=B2+B3"]] });
    expect(book.get("Sheet1", "E1")).toBe("=B2+B3");
  });
});

describe("notes", () => {
  it("attaches notes in the same call", async () => {
    const result = await call("write_range", { address: "E1:F1", values: [[1, 2]], notes: { F1: "Source: test" } });
    expect(result).toMatchObject({ ok: true, data: { notesWritten: 1 } });
    expect(sheet.noteTexts.get("F1")).toBe("Source: test");
  });

  it("updates an existing note instead of failing", async () => {
    sheet.noteTexts.set("F1", "old");
    await call("write_range", { address: "E1:F1", values: [[1, 2]], notes: { F1: "new" } });
    expect(sheet.noteTexts.get("F1")).toBe("new");
  });

  it("refuses before writing anything when notes are unsupported", async () => {
    book.supportsNotes = false;
    expectError(await call("write_range", { address: "E1", values: [[1]], notes: { E1: "x" } }), "unsupported");
    expect(book.get("Sheet1", "E1")).toBeNull();
  });

  it("rejects a note outside the written range", async () => {
    expectError(await call("write_range", { address: "E1", values: [[1]], notes: { Z9: "x" } }), "invalid_arguments");
    expect(book.get("Sheet1", "E1")).toBeNull();
  });
});

describe("set_formula fill", () => {
  it("fills a range and shifts relative references", async () => {
    const result = await call("set_formula", { address: "D2:D3", formula: "B2*C2" });
    expect(result.ok).toBe(true);
    expect(book.get("Sheet1", "D2")).toBe("=B2*C2");
    expect(book.get("Sheet1", "D3")).toBe("=B3*C3");
  });

  it("keeps absolute references fixed", async () => {
    await call("set_formula", { address: "D2:D3", formula: "B2*$C$2" });
    expect(book.get("Sheet1", "D3")).toBe("=B3*$C$2");
  });
});

describe("clear_range", () => {
  it("clears contents but keeps formatting by default", async () => {
    await call("format_range", { address: "A2", format: { bold: true } });
    await call("clear_range", { address: "A2:B2" });
    expect(book.get("Sheet1", "A2")).toBeNull();
    expect(sheet.formatLog.some((e) => e.props["cleared"])).toBe(false);
  });

  it("supports whole-row references and clearType all", async () => {
    const result = await call("clear_range", { address: "3:3", clearType: "all" });
    expect(result.ok).toBe(true);
    expect(book.get("Sheet1", "A3")).toBeNull();
    expect(sheet.formatLog.some((e) => e.props["cleared"] === "All")).toBe(true);
  });
});

describe("search_data", () => {
  beforeEach(() => {
    book.addSheet("Other");
    book.fill("Other", [["apple pie"], ["Apple"]]);
  });

  it("searches every visible sheet, case-insensitively by default", async () => {
    const result = await call("search_data", { searchTerm: "APPLE" });
    expect(result.ok && (result.data as any).matches).toEqual([
      { sheet: "Sheet1", address: "A2", value: "Apple" },
      { sheet: "Other", address: "A1", value: "apple pie" },
      { sheet: "Other", address: "A2", value: "Apple" },
    ]);
  });

  it("honours matchCase and wholeCell", async () => {
    const result = await call("search_data", { searchTerm: "Apple", matchCase: true, wholeCell: true });
    expect(result.ok && (result.data as any).matches.map((m: any) => `${m.sheet}!${m.address}`)).toEqual(["Sheet1!A2", "Other!A2"]);
  });

  it("pages results with nextOffset", async () => {
    const first = await call("search_data", { searchTerm: "apple", limit: 2 });
    expect(first.ok && (first.data as any)).toMatchObject({ returned: 2, hasMore: true, nextOffset: 2 });
    const second = await call("search_data", { searchTerm: "apple", limit: 2, offset: 2 });
    expect(second.ok && (second.data as any)).toMatchObject({ returned: 1, hasMore: false, nextOffset: null });
  });

  it("limits the search to a range", async () => {
    const result = await call("search_data", { searchTerm: "apple", sheet: { name: "Other" }, address: "A2:A2" });
    expect(result.ok && (result.data as any).matches).toHaveLength(1);
  });
});

describe("formula feedback", () => {
  it("set_formula reports what the formulas evaluated to", async () => {
    const result = await call("set_formula", { address: "D2:D3", formula: "B2*C2" });
    expect(result.ok && (result.data as any)).toMatchObject({ formulaResults: { D2: 4.5, D3: 4 } });
    expect(result.ok && (result.data as any).messages).toBeUndefined();
  });

  it("write_range explains that '=' strings became formulas", async () => {
    const result = await call("write_range", { address: "E1", values: [["=B2+B3"]] });
    expect(result.ok && (result.data as any)).toMatchObject({ formulaResults: { E1: 5 } });
    expect(result.ok && (result.data as any).messages[0]).toMatch(/stored as formulas/);
  });

  it("write_range of plain values adds no formula report", async () => {
    const result = await call("write_range", { address: "E1", values: [["a"]] });
    expect(result.ok && (result.data as any).formulaResults).toBeUndefined();
  });
});

describe("read_range includeText", () => {
  it("returns displayed text parallel to the values", async () => {
    const result = await call("read_range", { address: "B2:C2", includeText: true });
    expect(result.ok && (result.data as any)).toMatchObject({ values: [[3, 1.5]], text: [["3", "1.5"]] });
  });

  it("omits text by default", async () => {
    const result = await call("read_range", { address: "B2:C2" });
    expect(result.ok && (result.data as any).text).toBeUndefined();
  });
});

describe("find_errors", () => {
  it("lists error cells with their formulas and counts per type", async () => {
    book.fill("Sheet1", [["#REF!", "#N/A", "ok"]], "E1");
    const result = await call("find_errors", {});
    expect(result.ok && (result.data as any)).toMatchObject({
      totalFound: 2,
      hasMore: false,
      byType: { "#REF!": 1, "#N/A": 1 },
      errors: [{ sheet: "Sheet1", address: "E1", error: "#REF!" }, { sheet: "Sheet1", address: "F1", error: "#N/A" }],
    });
  });

  it("reports a clean sheet", async () => {
    const result = await call("find_errors", { sheet: { name: "Sheet1" } });
    expect(result.ok && (result.data as any)).toMatchObject({ totalFound: 0, errors: [] });
  });

  it("caps the listing but keeps the total", async () => {
    book.fill("Sheet1", [["#REF!", "#REF!", "#REF!"]], "E1");
    const result = await call("find_errors", { limit: 1 });
    expect(result.ok && (result.data as any)).toMatchObject({ totalFound: 3, returned: 1, hasMore: true });
  });
});

function resultAddresses(result: Awaited<ReturnType<typeof call>>): string[] {
  return result.ok ? (result.data as any).matches.map((m: any) => m.address) : [];
}

describe("search_data formulas and regex", () => {
  beforeEach(async () => {
    await call("set_formula", { address: "D2:D3", formula: "B2*C2" });
  });

  it("matches formula text and returns the formula", async () => {
    const result = await call("search_data", { searchTerm: "*C", matchFormulas: true });
    expect(result.ok && (result.data as any).matches).toEqual([
      { sheet: "Sheet1", address: "D2", value: 4.5, formula: "=B2*C2" },
      { sheet: "Sheet1", address: "D3", value: 4, formula: "=B3*C3" },
    ]);
  });

  it("does not look inside formulas unless asked", async () => {
    const result = await call("search_data", { searchTerm: "*C" });
    expect(result.ok && (result.data as any).matches).toEqual([]);
  });

  it("supports regular expressions, honouring wholeCell", async () => {
    const found = await call("search_data", { searchTerm: "^(Apple|Pear)$", useRegex: true, wholeCell: true });
    expect(resultAddresses(found)).toEqual(["A2", "A3"]);
    const partial = await call("search_data", { searchTerm: "^Ap", useRegex: true });
    expect(resultAddresses(partial)).toEqual(["A2"]);
  });

  it("rejects an invalid regular expression", async () => {
    expectError(await call("search_data", { searchTerm: "(", useRegex: true }), "invalid_arguments");
  });
});

describe("check_formula_consistency", () => {
  beforeEach(() => {
    const calc = book.addSheet("Calc");
    book.fill("Calc", [["Q", "P", "Total"], [1, 2, null], [2, 3, null], [3, 4, null], [4, 5, 99], [5, 6, null], [6, 7, null]]);
    const formulas: Record<number, string> = { 1: "=A2*B2", 2: "=A3*B3", 3: "=A4*B4", 5: "=A6+B6", 6: "=A7*B7" };
    for (const [row, f] of Object.entries(formulas)) calc.setCell(Number(row), 2, { v: null, f });
    calc.setCell(7, 2, { v: null, f: "=SUM(C2:C7)" });
  });

  it("flags a different formula and a hard-coded value but not the total", async () => {
    const result = await call("check_formula_consistency", { sheet: { name: "Calc" } });
    expect(result.ok && (result.data as any)).toMatchObject({
      totalIssues: 2,
      byKind: { different_formula: 1, hardcoded_value: 1 },
      linesChecked: 1,
      issues: [
        { address: "C5", kind: "hardcoded_value", value: 99, expected: { address: "C2", formula: "=A2*B2" } },
        { address: "C6", kind: "different_formula", formula: "=A6+B6", expected: { address: "C2", formula: "=A2*B2" } },
      ],
    });
  });

  it("reports a clean sheet", async () => {
    const result = await call("check_formula_consistency", { sheet: { name: "Sheet1" } });
    expect(result.ok && (result.data as any)).toMatchObject({ totalIssues: 0, issues: [] });
  });
});

describe("relativePattern", () => {
  it("makes copied formulas compare equal and keeps absolute refs and text intact", async () => {
    const { relativePattern } = await import("../addin/taskpane/executor/audit.js");
    expect(relativePattern("=B2*C2", 1, 3)).toBe(relativePattern("=B3*C3", 2, 3));
    expect(relativePattern("=B2*$C$1", 1, 3)).toBe(relativePattern("=B3*$C$1", 2, 3));
    expect(relativePattern("=B2*$C$1", 1, 3)).not.toBe(relativePattern("=B3*$C$2", 2, 3));
    expect(relativePattern('=IF(A1="A1",LOG10(A1),Sheet1!A1)', 0, 5)).toBe('=IF(R[0]C[-5]="A1",LOG10(R[0]C[-5]),Sheet1!R[0]C[-5])');
  });
});

describe("find_replace", () => {
  it("previews without changing anything", async () => {
    const result = await call("find_replace", { find: "apple", replace: "Mango" });
    expect(result.ok && (result.data as any)).toMatchObject({
      applied: false,
      totalMatches: 1,
      changes: [{ sheet: "Sheet1", address: "A2", before: "Apple", after: "Mango" }],
    });
    expect(book.get("Sheet1", "A2")).toBe("Apple");
  });

  it("applies the change and skips numbers", async () => {
    const result = await call("find_replace", { find: "a", replace: "4", apply: true });
    expect(result.ok && (result.data as any)).toMatchObject({ applied: true, replaced: 2 });
    expect(book.get("Sheet1", "A2")).toBe("4pple");
    expect(book.get("Sheet1", "B2")).toBe(3);
  });

  it("honours matchCase and wholeCell", async () => {
    const miss = await call("find_replace", { find: "apple", replace: "x", matchCase: true });
    expect(miss.ok && (miss.data as any).totalMatches).toBe(0);
    const whole = await call("find_replace", { find: "App", replace: "x", wholeCell: true });
    expect(whole.ok && (whole.data as any).totalMatches).toBe(0);
  });

  it("treats the replacement literally", async () => {
    await call("find_replace", { find: "Apple", replace: "$&-$1", apply: true });
    expect(book.get("Sheet1", "A2")).toBe("$&-$1");
  });

  it("keeps replaced text that starts with = as text", async () => {
    await call("find_replace", { find: "Apple", replace: "=x", apply: true });
    expect(book.get("Sheet1", "A2")).toBe("'=x");
  });

  it("edits formula text only when inFormulas is set", async () => {
    await call("set_formula", { address: "D2", formula: "B2*C2" });
    const untouched = await call("find_replace", { find: "B2", replace: "B3", apply: true });
    expect(untouched.ok && (untouched.data as any).replaced).toBe(0);
    const result = await call("find_replace", { find: "B2", replace: "B3", inFormulas: true, apply: true });
    expect(result.ok && (result.data as any).replaced).toBe(1);
    expect(book.get("Sheet1", "D2")).toBe("=B3*C2");
  });

  it("refuses to apply more than the cell limit", async () => {
    book.fill("Sheet1", Array.from({ length: 2_001 }, () => ["x"]), "H1");
    expectError(await call("find_replace", { find: "x", replace: "y", apply: true }), "size_limit_exceeded");
    expect(book.get("Sheet1", "H1")).toBe("x");
  });
});

describe("merge_range default", () => {
  it("merges when operation is omitted", async () => {
    const result = await call("merge_range", { address: "G1:H1" });
    expect(result.ok && (result.data as any).operation).toBe("merge");
  });
});

describe("get_sheets tables", () => {
  it("includes an empty tables list per sheet", async () => {
    const result = await call("get_sheets", {});
    expect(result.ok && (result.data as any).sheets[0].tables).toEqual([]);
  });
});

describe("format_range", () => {
  it("applies a number format to a multi-cell range", async () => {
    const result = await call("format_range", { address: "B2:C3", format: { numberFormat: "#,##0.00" } });
    expect(result.ok).toBe(true);
  });

  it("draws borders on a single cell without inside edges", async () => {
    await call("format_range", { address: "A1", format: { borders: "all" } });
    const keys = sheet.formatLog.flatMap((e) => Object.keys(e.props));
    expect(keys.some((k) => k.includes("EdgeTop"))).toBe(true);
    expect(keys.some((k) => k.includes("Inside"))).toBe(false);
  });

  it("adds inside edges for a block", async () => {
    await call("format_range", { address: "A1:C3", format: { borders: "all", autofitColumns: true } });
    const keys = sheet.formatLog.flatMap((e) => Object.keys(e.props));
    expect(keys.some((k) => k.includes("InsideHorizontal"))).toBe(true);
    expect(keys).toContain("autofitColumns");
  });
});

describe("modify_structure", () => {
  it("inserts, deletes and hides rows", async () => {
    await call("modify_structure", { operation: "insert", dimension: "rows", reference: "2", count: 1 });
    expect(book.get("Sheet1", "A3")).toBe("Apple");
    await call("modify_structure", { operation: "delete", dimension: "rows", reference: "2" });
    expect(book.get("Sheet1", "A2")).toBe("Apple");
    await call("modify_structure", { operation: "hide", dimension: "rows", reference: "2", count: 2 });
    expect([...sheet.hiddenRows]).toEqual([2, 3]);
    await call("modify_structure", { operation: "unhide", dimension: "rows", reference: "2", count: 2 });
    expect(sheet.hiddenRows.size).toBe(0);
  });

  it("inserts columns after a reference", async () => {
    await call("modify_structure", { operation: "insert", dimension: "columns", reference: "A", position: "after" });
    expect(book.get("Sheet1", "C1")).toBe("Qty");
    expect(book.get("Sheet1", "A1")).toBe("Item");
  });

  it("freezes and unfreezes panes", async () => {
    await call("modify_structure", { operation: "freeze", dimension: "rows", count: 1 });
    expect(sheet.frozen).toEqual({ rows: 1, columns: 0 });
    await call("modify_structure", { operation: "unfreeze" });
    expect(sheet.frozen).toEqual({ rows: 0, columns: 0 });
  });

  it("rejects inconsistent arguments at validation", () => {
    expect(() => parseToolInput("modify_structure", { operation: "insert", dimension: "rows", reference: "C" })).toThrow();
    expect(() => parseToolInput("modify_structure", { operation: "delete", dimension: "rows" })).toThrow();
    expect(() => parseToolInput("modify_structure", { operation: "freeze", dimension: "rows", reference: "2" })).toThrow();
  });
});

describe("shapes", () => {
  const box = (name: string, anchor: string, extra: Record<string, unknown> = {}) =>
    call("add_shape", { name, type: "FlowChartProcess", anchor, text: name, ...extra });

  it("draws a shape over a cell range with text and style", async () => {
    const result = await box("start", "M10:M11", { fillColor: "#DDEBF7", bold: true });
    expect(result).toMatchObject({ ok: true, data: { name: "start", type: "FlowChartProcess", text: "start" } });
    const shape = sheet.shapeList[0]!;
    expect(shape.left).toBe(12 * 48 + 2);
    expect(shape.top).toBe(9 * 15 + 2);
    expect(shape.width).toBe(48 - 4);
    expect(shape.fillColor).toBe("#DDEBF7");
    expect(shape.textFont["bold"]).toBe(true);
  });

  it("accepts explicit points and requires a complete placement", async () => {
    expect((await call("add_shape", { type: "Rectangle", left: 10, top: 20, width: 100, height: 40 })).ok).toBe(true);
    expect(() => parseToolInput("add_shape", { type: "Rectangle", left: 10 })).toThrow(/anchor/);
  });

  it("draws text boxes", async () => {
    const result = await call("add_shape", { name: "yes", type: "TextBox", text: "Yes", anchor: "N12" });
    expect(result).toMatchObject({ ok: true, data: { type: "TextBox", text: "Yes" } });
  });

  it("refuses duplicate names", async () => {
    await box("a", "M10");
    expectError(await box("a", "M12"), "invalid_arguments");
    expect(sheet.shapeList).toHaveLength(1);
  });

  it("connects shapes top-down with the facing sides and an arrowhead", async () => {
    await box("a", "M10:M11");
    await box("b", "M13:M14");
    const result = await call("add_connector", { from: "a", to: "b" });
    expect(result).toMatchObject({ ok: true, data: { fromSide: "bottom", toSide: "top", fromSite: 2, toSite: 0 } });
    const line = sheet.shapeList[2]!;
    expect(line.connections.begin?.shape.name).toBe("a");
    expect(line.connections.end?.shape.name).toBe("b");
    expect(line.lineProps["endArrowheadStyle"]).toBe("Triangle");
    expect(line.lineProps["connectorType"]).toBe("Elbow");
  });

  it("connects a decision branch sideways and handles the ellipse's eight sites", async () => {
    await call("add_shape", { name: "d", type: "FlowChartDecision", anchor: "M10:M12" });
    await call("add_shape", { name: "e", type: "Ellipse", anchor: "K10:K12" });
    const result = await call("add_connector", { from: "d", to: "e", fromSide: "left" });
    expect(result).toMatchObject({ ok: true, data: { fromSide: "left", toSide: "right", fromSite: 1, toSite: 6 } });
  });

  it("fails clearly for a missing shape and lists what exists", async () => {
    await box("a", "M10");
    const error = expectError(await call("add_connector", { from: "a", to: "ghost" }), "invalid_arguments");
    expect(error?.message).toMatch(/target shape 'ghost' not found/);
    expect((error?.detail as any).available).toEqual(["a"]);
  });

  it("removes the half-made connector when Excel rejects the site", async () => {
    await box("a", "M10");
    await box("b", "M13");
    expectError(await call("add_connector", { from: "a", to: "b", fromSite: 40 }), "invalid_arguments");
    expect(sheet.shapeList.filter((s) => !s.deleted).map((s) => s.name)).toEqual(["a", "b"]);
  });

  it("asks for explicit sites on shape types whose layout is unknown", async () => {
    sheet.shapes.addGeometricShape("Hexagon").name = "hex";
    await box("a", "M10");
    expectError(await call("add_connector", { from: "a", to: "hex" }), "invalid_arguments");
  });

  it("lists, updates and deletes shapes", async () => {
    await box("a", "M10");
    await call("add_shape", { name: "note", type: "TextBox", text: "hi", anchor: "O10" });
    const listed = await call("list_shapes", {});
    expect(listed.ok && (listed.data as any).shapes.map((s: any) => [s.name, s.type, s.text])).toEqual([
      ["a", "FlowChartProcess", "a"],
      ["note", "TextBox", "hi"],
    ]);

    const restyled = await call("update_shape", { name: "a", fillColor: "#FFFF00" });
    expect(restyled).toMatchObject({ ok: true, data: { name: "a", text: "a" } }); // keeps reporting the existing text
    const updated = await call("update_shape", { name: "a", newName: "first", text: "First", fillColor: "#FFFF00", anchor: "M20" });
    expect(updated).toMatchObject({ ok: true, data: { name: "first" } });
    expect(sheet.shapeList[0]!.top).toBe(19 * 15 + 2);

    const removed = await call("delete_shapes", { names: ["first", "missing"] });
    expect(removed).toMatchObject({ ok: true, data: { deleted: ["first"], notFound: ["missing"] } });
  });

  it("rounds fractional line weights to whole points, as Excel requires", async () => {
    await box("a", "M10");
    const result = await call("update_shape", { name: "a", lineColor: "#6AA84F", lineWeight: 1.5, fontSize: 12, bold: true });
    expect(result).toMatchObject({ ok: true, data: { name: "a", warnings: ["line weight 1.5 rounded to 2 pt (Excel needs whole points)"] } });
    expect(sheet.shapeList[0]!.lineFormat["weight"]).toBe(2);
    expect(sheet.shapeList[0]!.lineFormat["color"]).toBe("#6AA84F");

    const sub = await call("add_shape", { name: "thin", type: "Rectangle", anchor: "O10", lineWeight: 0.25 });
    expect(sub).toMatchObject({ ok: true, data: { warnings: [expect.stringContaining("rounded to 1 pt")] } });
    expect((await call("update_shape", { name: "a", lineWeight: 3 })).ok).toBe(true);
  });

  it("rounds connector weights too", async () => {
    await box("a", "M10:M11");
    await box("b", "M13:M14");
    const result = await call("add_connector", { from: "a", to: "b", weight: 1.25, color: "#333333" });
    expect(result).toMatchObject({ ok: true, data: { warnings: [expect.stringContaining("1.25 rounded to 1 pt")] } });
    expect(sheet.shapeList[2]!.lineFormat["weight"]).toBe(1);
  });

  describe("updating connectors (type Line)", () => {
    beforeEach(async () => {
      await box("a", "M10:M11");
      await box("b", "M13:M14");
      await call("add_connector", { name: "c1", from: "a", to: "b" });
    });

    it("changes outline and name of a connector without touching its body (which Excel forbids)", async () => {
      const result = await call("update_shape", { name: "c1", lineColor: "#112233", lineWeight: 2, newName: "c2" });
      expect(result).toMatchObject({ ok: true, data: { name: "c2", type: "Line", text: null } });
      const line = sheet.shapeList[2]!;
      expect(line.lineFormat["color"]).toBe("#112233");
      expect(line.lineFormat["weight"]).toBe(2);
    });

    it("explains which properties do not apply to a connector", async () => {
      const error = expectError(await call("update_shape", { name: "c1", fillColor: "#FFFFFF", text: "x", fontName: "Arial" }), "invalid_arguments");
      expect(error?.message).toMatch(/line\/connector.*text, fillColor, fontName|line\/connector.*fillColor/);
      expect(error?.message).toMatch(/delete_shapes/);
    });

    it("lists connectors alongside shapes and can connect to a text box only with explicit sites", async () => {
      await call("add_shape", { name: "lbl", type: "TextBox", text: "Yes", anchor: "P10" });
      const listed = await call("list_shapes", {});
      expect(listed.ok && (listed.data as any).shapes.map((s: any) => s.type)).toEqual(["FlowChartProcess", "FlowChartProcess", "Line", "TextBox"]);
      expectError(await call("add_connector", { from: "a", to: "lbl" }), "invalid_arguments");
    });
  });

  it("reports an unsupported Excel version instead of failing obscurely", async () => {
    book.supportsShapes = false;
    expectError(await box("a", "M10"), "unsupported");
  });
});

describe("lines and dashed connectors", () => {
  it("draws a free line between two cells with dash and arrowhead", async () => {
    const result = await call("add_line", { name: "legend_popup", from: "D8", to: "G8", dash: "longDash", arrow: "end", color: "#000000", weight: 1 });
    expect(result).toMatchObject({ ok: true, data: { name: "legend_popup" } });
    const line = sheet.shapeList[0]!;
    expect(line.lineFormat["dashStyle"]).toBe("LongDash");
    expect(line.lineProps["endArrowheadStyle"]).toBe("Triangle");
    expect(line.width).toBe(3 * 48);
  });

  it("supports diamond line ends and refuses duplicate names", async () => {
    await call("add_line", { name: "end", from: "N8", to: "Q8", arrow: "diamond" });
    expect(sheet.shapeList[0]!.lineProps["endArrowheadStyle"]).toBe("Diamond");
    expectError(await call("add_line", { name: "end", from: "N9", to: "Q9" }), "invalid_arguments");
  });

  it("applies dash styles to connectors too", async () => {
    await call("add_shape", { name: "a", type: "Rectangle", anchor: "A1:B2" });
    await call("add_shape", { name: "b", type: "Rectangle", anchor: "A6:B7" });
    await call("add_connector", { from: "a", to: "b", dash: "dashDotDot", arrow: "diamond" });
    const line = sheet.shapeList[2]!;
    expect(line.lineFormat["dashStyle"]).toBe("DashDotDot");
    expect(line.lineProps["endArrowheadStyle"]).toBe("Diamond");
  });

  it("accepts a font name for shape text", async () => {
    await call("add_shape", { name: "p", type: "FlowChartProcess", anchor: "A1:C2", text: "x", fontName: "Arial", fontSize: 8 });
    expect(sheet.shapeList[0]!.textFont["name"]).toBe("Arial");
  });
});


describe("shape selection and flow inspection", () => {
  async function flow() {
    await call("add_shape", { name: "start", type: "FlowChartTerminator", text: "Start", anchor: "A5:B6" });
    await call("add_shape", { name: "step", type: "FlowChartProcess", text: "Process", anchor: "A9:B10" });
    await call("add_connector", { name: "next", from: "start", to: "step" });
  }

  it("identifies the active shape even when getSelectedRange would return stale cells", async () => {
    await flow();
    book.activeShape = sheet.shapeList[1]!;
    const result = await call("get_selection", {});
    expect(result).toMatchObject({ ok: true, data: { address: "", values: [], shapeSelectionStatus: "active", activeShape: { name: "step", text: "Process", id: book.activeShape.id } } });
    if (result.ok) expect(SelectionSnapshotSchema.parse(result.data).activeShape?.name).toBe("step");
  });

  it("reads an active connector's endpoints without accessing its text frame", async () => {
    await flow();
    book.invalidSelection = true;
    book.activeShape = sheet.shapeList[2]!;
    expect(await call("get_selection", {})).toMatchObject({ ok: true, data: { activeShape: { name: "next", connection: { from: { name: "start" }, to: { name: "step" }, endArrow: "Triangle" } } } });
  });

  it("explains unavailable shape identification on older hosts", async () => {
    book.supportsActiveShape = false;
    book.invalidSelection = true;
    expect(await call("get_selection", {})).toMatchObject({ ok: true, data: { shapeSelectionStatus: "unavailable", note: expect.stringContaining("ExcelApi 1.19") } });
  });

  it("reads a connected flow and preserves connector IDs, sites and arrows", async () => {
    await flow();
    const result = await call("list_shapes", {});
    expect(result).toMatchObject({ ok: true, data: { count: 3, hasMore: false, validation: { complete: true, nodeCount: 2, lineCount: 1, issues: [], components: [[sheet.shapeList[0]!.id, sheet.shapeList[1]!.id]] } } });
    if (result.ok) expect((result.data as any).shapes[2].connection).toMatchObject({ from: { name: "start", site: 2 }, to: { name: "step", site: 0 }, endArrow: "Triangle" });
  });

  it("exposes nested groups with parent IDs", async () => {
    await flow();
    const nested = new MockShape("Group", "");
    nested.name = "inner";
    nested.children = sheet.shapeList.splice(0);
    const outer = new MockShape("Group", "");
    outer.name = "outer";
    outer.children = [nested];
    sheet.shapeList.push(outer);
    expect(await call("list_shapes", {})).toMatchObject({ ok: true, data: { count: 5, shapes: [
      { name: "outer" }, { name: "inner", parentGroupId: outer.id },
      { name: "start", parentGroupId: nested.id }, { name: "step", parentGroupId: nested.id },
      { name: "next", parentGroupId: nested.id, connection: { from: { name: "start" }, to: { name: "step" } } },
    ] } });
  });

  it("reports unattached lines and isolated shapes without asserting business validity", async () => {
    await flow();
    sheet.shapeList[2]!.connections.end = undefined;
    expect(await call("list_shapes", {})).toMatchObject({ ok: true, data: { validation: { issues: [
      { code: "unattached_line", shapes: ["next"] }, { code: "isolated_shape", shapes: ["step"] },
    ] } } });
  });

  it("paginates beyond 200 shapes and marks partial checks incomplete", async () => {
    for (let i = 0; i < 205; i++) sheet.shapes.addGeometricShape("Rectangle").name = `n${i}`;
    const first = await call("list_shapes", {});
    expect(first).toMatchObject({ ok: true, data: { count: 205, returned: 200, nextOffset: 200, hasMore: true, validation: { complete: false, issues: [] } } });
    const last = await call("list_shapes", { offset: 200 });
    expect(last).toMatchObject({ ok: true, data: { returned: 5, nextOffset: null, hasMore: false, validation: { complete: false } } });
    expect(await call("list_shapes", { limit: 1000 })).toMatchObject({ ok: true, data: { returned: 205, validation: { complete: true } } });
  });
});


describe("validate_flowchart tool", () => {
  it("runs the read-only validator through the registry and live executor", async () => {
    await call("add_shape", { name: "a", type: "Rectangle", anchor: "A5:B6" });
    await call("add_shape", { name: "b", type: "Rectangle", anchor: "A9:B10" });
    await call("add_connector", { name: "link", from: "a", to: "b", connectorType: "Straight" });
    const before = sheet.shapeList.map(s => ({ name: s.name, left: s.left, top: s.top }));
    expect(await call("validate_flowchart", {})).toMatchObject({ ok: true, data: {
      status: "no_issues_in_checked_geometry", scanned: 3, scanTruncated: false,
      geometry: { complete: true, checkedConnectors: 1, issues: [] },
    } });
    expect(sheet.shapeList.map(s => ({ name: s.name, left: s.left, top: s.top }))).toEqual(before);
  });

  it("does not certify routed connectors or truncated scans", async () => {
    await call("add_shape", { name: "a", type: "Rectangle", anchor: "A5:B6" });
    await call("add_shape", { name: "b", type: "Rectangle", anchor: "A9:B10" });
    await call("add_connector", { from: "a", to: "b", connectorType: "Elbow" });
    expect(await call("validate_flowchart", {})).toMatchObject({ ok: true, data: { status: "needs_visual_review", geometry: { complete: false } } });
    expect(await call("validate_flowchart", { maxShapes: 1 })).toMatchObject({ ok: true, data: { status: "needs_visual_review", scanTruncated: true } });
  });
});


describe("optional shape text failures", () => {
  async function sheetWithBrokenText(mode: "frame" | "range" | "hasTextSync" | "rangeSync") {
    await call("add_shape", { name: "good", type: "Rectangle", anchor: "A5:B6", text: "Readable step" });
    await call("add_shape", { name: "bad", type: "Rectangle", anchor: "A9:B10", text: "Unreachable text" });
    await call("add_shape", { name: "after", type: "Rectangle", anchor: "A13:B14", text: "Still readable" });
    await call("add_connector", { name: "link", from: "good", to: "bad", connectorType: "Straight" });
    await call("add_connector", { name: "next", from: "bad", to: "after", connectorType: "Straight" });
    sheet.shapeList[1]!.textReadFailure = mode;
  }

  it("does not navigate TextRange on an empty shape", async () => {
    await call("add_shape", { name: "empty", type: "Rectangle", anchor: "A5:B6" });
    const empty = sheet.shapeList[0]!;
    empty.emptyTextRangeFails = true;
    empty.textRangeAccesses = 0;
    expect(await call("list_shapes", {})).toMatchObject({ ok: true, data: {
      textReadComplete: true, readWarnings: [], shapes: [{ name: "empty", text: "", textReadStatus: "empty" }],
    } });
    expect(empty.textRangeAccesses).toBe(0);
  });

  it.each(["frame", "range", "hasTextSync", "rangeSync"] as const)("keeps the whole sheet and connector data when text fails at %s", async mode => {
    await sheetWithBrokenText(mode);
    const result = await call("list_shapes", {});
    expect(result).toMatchObject({ ok: true, data: {
      count: 5, returned: 5, textReadComplete: false,
      readWarnings: [{ name: "bad", property: "text" }],
      shapes: [
        { name: "good", text: "Readable step", textReadStatus: "read" },
        { name: "bad", text: null, textReadStatus: "unavailable", left: expect.any(Number) },
        { name: "after", text: "Still readable", textReadStatus: "read" },
        { name: "link", textReadStatus: "not_applicable", connection: { from: { name: "good" }, to: { name: "bad" } } },
        { name: "next", connection: { from: { name: "bad" }, to: { name: "after" } } },
      ],
    } });
    expect(sheet.shapeList).toHaveLength(5);
    expect(sheet.shapeList.every(shape => !shape.deleted)).toBe(true);
    // A rejected optional batch must not poison the next tool request.
    expect(await call("get_sheets", {})).toMatchObject({ ok: true });
  });

  it("keeps distinct diagnostics for two unreadable shapes with the same name", async () => {
    await sheetWithBrokenText("rangeSync");
    sheet.shapeList[0]!.name = "Text Box 1";
    sheet.shapeList[0]!.textReadFailure = "rangeSync";
    sheet.shapeList[1]!.name = "Text Box 1";
    const result = await call("list_shapes", {});
    expect(result).toMatchObject({ ok: true, data: {
      count: 5, textReadComplete: false,
      shapes: [{ textReadStatus: "unavailable" }, { textReadStatus: "unavailable" }, { text: "Still readable" }, {}, {}],
      readWarnings: [{ id: sheet.shapeList[0]!.id, name: "Text Box 1" }, { id: sheet.shapeList[1]!.id, name: "Text Box 1" }],
    } });
  });

  it("continues geometric validation and reports the unreadable text separately", async () => {
    await sheetWithBrokenText("rangeSync");
    expect(await call("validate_flowchart", {})).toMatchObject({ ok: true, data: {
      scanned: 5, status: "no_issues_in_checked_geometry", textReadComplete: false,
      readWarnings: [{ name: "bad", property: "text" }],
      geometry: { complete: true, checkedConnectors: 2, issues: [] },
    } });
  });

  it("can identify a selected shape even when its text is unreadable", async () => {
    await sheetWithBrokenText("range");
    book.activeShape = sheet.shapeList[1]!;
    const result = await call("get_selection", {});
    expect(result).toMatchObject({ ok: true, data: { activeShape: { name: "bad", text: null, textReadStatus: "unavailable" } } });
    if (result.ok) expect(SelectionSnapshotSchema.parse(result.data).activeShape?.readWarnings).toHaveLength(1);
  });
});


describe("editing grouped shapes", () => {
  it("updates the selected nested child directly without changing its siblings or group", async () => {
    const child = new MockShape("GeometricShape", "Rectangle"); child.name = "フローチャート: 処理 33";
    const sibling = new MockShape("GeometricShape", "Rectangle"); sibling.name = "Sibling";
    const inner = new MockShape("Group", ""); inner.name = "Inner"; inner.children = [child, sibling];
    const outer = new MockShape("Group", ""); outer.name = "Group 62"; outer.children = [inner];
    sheet.shapeList.push(outer);
    book.activeShape = child;
    expect(await call("update_shape", { sheet: { name: "Sheet1" }, name: child.name, fillColor: "#FF0000" })).toMatchObject({ ok: true, data: { name: child.name } });
    expect(child.fillColor).toBe("#FF0000");
    expect(sibling.fillColor).toBeNull();
    expect(outer.fillColor).toBeNull();
    expect(book.activeShape).toBe(child);
  });

  it("updates grouped connectors using lineColor", async () => {
    const line = new MockShape("Line", ""); line.name = "Selected connector";
    const group = new MockShape("Group", ""); group.name = "Group"; group.children = [line];
    sheet.shapeList.push(group);
    expect(await call("update_shape", { name: line.name, lineColor: "#FF0000" })).toMatchObject({ ok: true });
    expect(line.lineFormat.color).toBe("#FF0000");
  });

  it("rejects ambiguous names inside different groups before writing", async () => {
    const children = [new MockShape("GeometricShape", "Rectangle"), new MockShape("GeometricShape", "Rectangle")];
    for (const [index, child] of children.entries()) {
      child.name = "Duplicate";
      const group = new MockShape("Group", ""); group.name = `Group ${index}`; group.children = [child];
      sheet.shapeList.push(group);
    }
    const result = await call("update_shape", { name: "Duplicate", fillColor: "#FF0000" });
    expectError(result, "invalid_arguments");
    expect(children.map(child => child.fillColor)).toEqual([null, null]);
  });
});

describe("a selected chart", () => {
  // Reading shape-only properties (rotation, geometry) off a selected chart crashed Excel, so charts go through the chart API.
  it("is labelled and described through the chart API", async () => {
    book.activeChart = { name: "MsgPerDay", left: 15, top: 230.25, width: 354, height: 240 };
    expect(await getSelectionLabel()).toEqual({ label: "MsgPerDay", title: "Selected chart: MsgPerDay" });
    expect(await call("get_selection", {})).toMatchObject({
      ok: true,
      data: { shapeSelectionStatus: "active", activeShape: { name: "MsgPerDay", type: "Chart", left: 15, top: 230.25, width: 354, height: 240 } },
    });
  });

  it("leaves other unsupported active objects unidentified instead of reading their geometry", async () => {
    await call("add_shape", { name: "odd object", type: "Rectangle", anchor: "A5:B6" });
    const odd = sheet.shapeList[0]!;
    (odd as { type: string }).type = "Unsupported";
    book.activeShape = odd;
    book.invalidSelection = true; // Excel refuses getSelectedRange while an object is selected
    const result = await call("get_selection", {});
    expect(result).toMatchObject({ ok: true, data: { shapeSelectionStatus: "unavailable" } });
    expect((result as { data: Record<string, unknown> }).data.activeShape).toBeUndefined();
  });
});

describe("lightweight selection labels", () => {
  it("reads connector names without accessing their text or connection metadata", async () => {
    await call("add_line", { name: "Connector_A", from: "A5", to: "C5" });
    await call("add_line", { name: "Connector_B", from: "A7", to: "C7" });
    book.activeShape = sheet.shapeList[0]!;
    expect(await getSelectionLabel()).toMatchObject({ label: "Connector_A", title: expect.stringContaining("line / connector") });
    book.activeShape = sheet.shapeList[1]!;
    expect(await getSelectionLabel()).toMatchObject({ label: "Connector_B" });
    book.activeShape = null;
    expect(await getSelectionLabel()).toMatchObject({ label: "Sheet1!A2:C2 · 3 cells" });
  });

  it("ignores broken text frames while identifying the selected shape", async () => {
    await call("add_shape", { name: "broken text", type: "Rectangle", anchor: "A5:B6" });
    book.activeShape = sheet.shapeList[0]!;
    book.activeShape.textReadFailure = "frame";
    expect(await getSelectionLabel()).toMatchObject({ label: "broken text" });
  });
});
