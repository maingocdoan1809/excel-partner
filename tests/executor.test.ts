import { beforeEach, describe, expect, it } from "vitest";
import { executeTool } from "../addin/taskpane/executor/index.js";
import { parseToolInput } from "../protocol/src/index.js";
import { MockWorkbook, MockSheet, installOfficeMock } from "./support/office-mock.js";

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
