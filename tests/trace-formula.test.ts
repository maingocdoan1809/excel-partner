import { beforeEach, describe, expect, it } from "vitest";
import { CONNECTED_READ_TOOLS, TraceFormulaInputSchema, parseToolInput, splitSheetAddress } from "../protocol/src/index.js";
import { executeTool } from "../frontend/taskpane/executor/index.js";
import { MockWorkbook, installOfficeMock } from "./support/office-mock.js";

let book: MockWorkbook;
beforeEach(() => {
  book = new MockWorkbook();
  book.addSheet("Calc");
  book.addSheet("Inputs");
  book.fill("Inputs", [[10, 20], [30, 40]]);
  book.fill("Calc", [[null, null, null], ["note", null, null]]);
  // fill() stores text as values; these are formulas.
  const calc = book.sheet("Calc");
  for (const [row, column, formula] of [[0, 0, "=Inputs!A1+Inputs!B1"], [0, 1, "=A1*2"], [0, 2, "=SUM(Inputs!A1:B2)"], [1, 1, "=B1+C1"]] as const) calc.setCell(row, column, { v: null, f: formula });
  book.selection = { sheet: "Calc", address: "A1" };
  installOfficeMock(book);
});

const call = (args: unknown) => executeTool("trace_formula", parseToolInput("trace_formula" as never, args));
const addresses = (result: any, direction?: string) => result.data.nodes.filter((node: any) => !direction || node.direction === direction).map((node: any) => `${node.level}:${node.address}`);

describe("splitSheetAddress", () => {
  it("separates the sheet from the range, unquoting names that need it", () => {
    expect(splitSheetAddress("Sheet1!A1:B2")).toEqual({ sheet: "Sheet1", range: "A1:B2" });
    expect(splitSheetAddress("'My Sheet'!$A$1")).toEqual({ sheet: "My Sheet", range: "$A$1" });
    expect(splitSheetAddress("'Bob''s'!C3")).toEqual({ sheet: "Bob's", range: "C3" });
    expect(splitSheetAddress("A1:B2")).toEqual({ sheet: null, range: "A1:B2" });
  });
});

describe("trace_formula", () => {
  it("lists what a formula reads, across sheets, with each cell's formula and value", async () => {
    const result = await call({ sheet: { name: "Calc" }, address: "A1" });
    expect(result).toMatchObject({ ok: true, data: { sheet: "Calc", address: "A1", formula: "=Inputs!A1+Inputs!B1", direction: "precedents", depth: 1, truncated: false } });
    expect(addresses(result)).toEqual(["1:Inputs!A1", "1:Inputs!B1"]);
    expect((result as any).data.nodes[0]).toMatchObject({ cells: 1, value: 10 });
  });

  it("follows the chain as many steps as asked, and stops at cells that read nothing", async () => {
    const one = await call({ sheet: { name: "Calc" }, address: "B2", depth: 1 });
    expect(addresses(one)).toEqual(["1:Calc!B1", "1:Calc!C1"]);
    const two = await call({ sheet: { name: "Calc" }, address: "B2", depth: 2 });
    expect(addresses(two)).toEqual(["1:Calc!B1", "1:Calc!C1", "2:Calc!A1", "2:Inputs!A1:B2"]);
    expect((two as any).data.nodes.find((node: any) => node.address === "Calc!B1")).toMatchObject({ formula: "=A1*2" });
    const three = await call({ sheet: { name: "Calc" }, address: "B2", depth: 3 });
    expect(addresses(three)).toContain("3:Inputs!A1");
  });

  it("lists what reads a cell, on this sheet and others", async () => {
    const result = await call({ sheet: { name: "Inputs" }, address: "A1", direction: "dependents" });
    expect(addresses(result)).toEqual(["1:Calc!A1", "1:Calc!C1"]);
    const deeper = await call({ sheet: { name: "Inputs" }, address: "A1", direction: "dependents", depth: 3 });
    // B2 is reached at step 2 (through C1) and again at step 3 (through B1): it is listed once, at the nearest step.
    expect(addresses(deeper)).toEqual(["1:Calc!A1", "1:Calc!C1", "2:Calc!B1", "2:Calc!B2"]);
  });

  it("can trace both ways at once, and says when a cell has no formula", async () => {
    const both = await call({ sheet: { name: "Calc" }, address: "B1", direction: "both" });
    expect(addresses(both, "precedents")).toEqual(["1:Calc!A1"]);
    expect(addresses(both, "dependents")).toEqual(["1:Calc!B2"]);
    const plain = await call({ sheet: { name: "Inputs" }, address: "A2" });
    expect((plain as any).data.nodes).toEqual([]);
    expect((plain as any).data.notes[0]).toContain("holds no formula");
  });

  it("does not list a cell twice when two routes lead to it", async () => {
    const result = await call({ sheet: { name: "Calc" }, address: "B2", depth: 3 });
    const names = addresses(result).map((entry: string) => entry.slice(entry.indexOf(":") + 1));
    expect(new Set(names).size).toBe(names.length);
  });

  it("stops at the limit and says so", async () => {
    const result = await call({ sheet: { name: "Calc" }, address: "B2", depth: 3, limit: 2 });
    expect((result as any).data).toMatchObject({ truncated: true });
    expect((result as any).data.nodes).toHaveLength(2);
    expect((result as any).data.notes.join(" ")).toContain("Stopped at 2");
  });

  it("takes one cell, with sensible defaults, and can be run in another open workbook", () => {
    expect(TraceFormulaInputSchema.parse({ address: "D10" })).toMatchObject({ direction: "precedents", depth: 1, limit: 100 });
    expect(() => TraceFormulaInputSchema.parse({ address: "A1:B2" })).toThrow();
    expect(() => TraceFormulaInputSchema.parse({ address: "A1", depth: 4 })).toThrow();
    expect(CONNECTED_READ_TOOLS).toContain("trace_formula");
  });
});
