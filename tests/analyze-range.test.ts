import { beforeEach, describe, expect, it } from "vitest";
import {
  AnalyzeError, AnalyzeRangeInputSchema, CONNECTED_READ_TOOLS, RangeAnalyzer, columnLetter, isDateFormat, isoToSerial, parseToolInput, resolveColumn,
  type AnalyzeColumn,
} from "../protocol/src/index.js";
import { executeTool } from "../frontend/taskpane/executor/index.js";
import { MockWorkbook, installOfficeMock } from "./support/office-mock.js";

const NAMES = ["Region", "Product", "Units", "Revenue", "Date"];
const columns = (dateLike = [false, false, false, false, true]): AnalyzeColumn[] => NAMES.map((name, index) => ({ index, letter: columnLetter(index), name, dateLike: dateLike[index]! }));
const D = (iso: string) => isoToSerial(iso)!;
const rows: unknown[][] = [
  ["North", "Pen", 10, 100, D("2025-01-10")],
  ["North", "Ink", 5, 80.5, D("2025-02-03")],
  ["south", "Pen", 7, 70, D("2025-02-20")],
  ["South", "Pad", "n/a", 30, D("2025-03-05")],
  ["East", "Pen", 3, "#DIV/0!", D("2025-03-31")],
  [null, null, null, null, null],
  ["East", "Ink", 2, 20, D("2025-04-01")],
];
const types = (data: unknown[][]) => data.map(row => row.map(value => value === null ? "Empty" : typeof value === "number" ? "Double" : /^#/.test(String(value)) ? "Error" : "String"));

function run(input: Record<string, unknown>, data = rows, dateLike?: boolean[]) {
  const parsed = AnalyzeRangeInputSchema.parse({ address: "A1:E8", ...input });
  const analyzer = new RangeAnalyzer(columns(dateLike), parsed);
  analyzer.add(data, types(data), 2);
  return analyzer.result() as any;
}

describe("analyze_range engine", () => {
  it("names columns by letters and recognises dates from the number format", () => {
    expect(["A", "Z", "AA", "AZ", "BA"].map((letter, i) => columnLetter([0, 25, 26, 51, 52][i]!))).toEqual(["A", "Z", "AA", "AZ", "BA"]);
    expect(["yyyy-mm-dd", "d/m/yy", "h:mm", "[$-409]d-mmm", "dd \"of\" mmm"].map(isDateFormat)).toEqual([true, true, true, true, true]);
    expect(["General", "0.00", "$#,##0.00", "0%", "@", "\"day\" 0"].map(isDateFormat)).toEqual([false, false, false, false, false, false]);
    expect(isoToSerial("2023-03-15")).toBe(45_000);
    expect(isoToSerial("2023-02-30")).toBeNull();
    expect(isoToSerial("March")).toBeNull();
  });

  it("finds a column by header in any case or by its letters, and explains what exists when it cannot", () => {
    const all = columns();
    expect(resolveColumn(all, " revenue ").name).toBe("Revenue");
    expect(resolveColumn(all, "c").name).toBe("Units");
    expect(() => resolveColumn(all, "Profit")).toThrow(/No column 'Profit'.*Region \(A\), Product \(B\)/);
    const twins = [...all, { index: 5, letter: "F", name: "Units", dateLike: false }];
    expect(() => resolveColumn(twins, "Units")).toThrow(/ambiguous.*C, F/);
  });

  it("groups and totals, merging 'North'/'north', ignoring text and errors in sums, and ordering by an aggregate", () => {
    const result = run({ mode: "query", groupBy: ["Region"], aggregate: [{ fn: "sum", column: "Revenue", as: "revenue" }, { fn: "sum", column: "Units", as: "units" }, { fn: "count" }] });
    expect(result.matched).toBe(6);
    expect(result.groups).toHaveLength(3);
    const byRegion = Object.fromEntries(result.groups.map((group: any) => [String(group.Region).toLowerCase(), group]));
    expect(byRegion.north).toMatchObject({ rows: 2, revenue: 180.5, units: 15, count: 2 });
    expect(byRegion.south).toMatchObject({ rows: 2, revenue: 100, units: 7 });
    expect(byRegion.east).toMatchObject({ rows: 2, revenue: 20, units: 5 });
    expect(result.groups[0].revenue).toBeGreaterThanOrEqual(result.groups[1].revenue);
    expect(result.notes.join(" ")).toContain("1 cells in the matching rows show an Excel error");
  });

  it("filters with several conditions, on text, numbers, dates given as YYYY-MM-DD, and blanks", () => {
    expect(run({ mode: "query", where: [{ column: "Product", op: "eq", value: "pen" }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 3 });
    expect(run({ mode: "query", where: [{ column: "Units", op: "gte", value: 5 }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 3 });
    expect(run({ mode: "query", where: [{ column: "Revenue", op: "between", value: 30, value2: 100 }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 4 });
    expect(run({ mode: "query", where: [{ column: "Date", op: "gte", value: "2025-02-01" }, { column: "Date", op: "lt", value: "2025-03-31" }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 3 });
    expect(run({ mode: "query", where: [{ column: "Region", op: "in", values: ["east", "NORTH"] }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 4 });
    expect(run({ mode: "query", where: [{ column: "Product", op: "contains", value: "in" }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 2 });
    expect(run({ mode: "query", where: [{ column: "Units", op: "blank" }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 0 });
    expect(run({ mode: "query", where: [{ column: "Product", op: "notContains", value: "pen" }], aggregate: [{ fn: "count" }] }).totals).toEqual({ count: 3 });
  });

  it("computes totals, average, min, max and distinct, and says how many non-numeric cells it ignored", () => {
    const result = run({ mode: "query", aggregate: [{ fn: "sum", column: "Units" }, { fn: "avg", column: "Units" }, { fn: "min", column: "Revenue" }, { fn: "max", column: "Revenue" }, { fn: "distinct", column: "Product" }] });
    expect(result.totals).toEqual({ sum_Units: 27, avg_Units: 5.4, min_Revenue: 20, max_Revenue: 100, distinct_Product: 3 });
    expect(result.notes.join(" ")).toContain("sum_Units: 1 non-numeric cells were ignored");
  });

  it("lists matching rows with their sheet row numbers, ordered and limited, only the columns asked for", () => {
    const result = run({ mode: "query", where: [{ column: "Product", op: "eq", value: "Pen" }], columns: ["Region", "Units"], orderBy: { by: "Units", direction: "desc" }, limit: 2 });
    expect(result.columns).toEqual(["Region", "Units"]);
    expect(result.rows).toEqual([["North", 10], ["south", 7]]);
    expect(result.rowNumbers).toEqual([2, 4]);
    expect(result).toMatchObject({ matched: 3, hasMore: true });
  });

  it("skips blank rows, and counts the cells it scanned", () => {
    expect(run({ mode: "query", aggregate: [{ fn: "count" }] }).rowsScanned).toBe(6);
  });

  it("profiles every column: type, blanks, errors, range, distinct values and the commonest ones", () => {
    const profile = run({ mode: "profile" }).columns as any[];
    const by = Object.fromEntries(profile.map(column => [column.name, column]));
    expect(by.Region).toMatchObject({ type: "text", filled: 6, distinct: 3 });
    expect(by.Region.top[0]).toMatchObject({ value: "North", count: 2 });
    expect(by.Units).toMatchObject({ type: "mixed", min: 2, max: 10, mix: { numbers: 5, text: 1, booleans: 0 } });
    expect(by.Revenue).toMatchObject({ errors: 1, min: 20, max: 100 });
    expect(by.Date).toMatchObject({ dateLike: true, type: "number" });
    expect(run({ mode: "profile", columns: ["Product"] }).columns).toHaveLength(1);
  });

  it("reports an unknown column, or an order it cannot apply, as a clear error", () => {
    expect(() => run({ mode: "query", where: [{ column: "Profit", op: "eq", value: 1 }] })).toThrow(AnalyzeError);
    expect(() => run({ mode: "query", groupBy: ["Region"], orderBy: { by: "Nope" } })).toThrow(/Cannot order by 'Nope'. Use one of: Region, rows, count/);
  });

  it("refuses conditions that lack what they compare with", () => {
    expect(() => parseToolInput("analyze_range" as never, { address: "A1:B3", mode: "query", where: [{ column: "A", op: "gt" }] })).toThrow();
    expect(() => parseToolInput("analyze_range" as never, { address: "A1:B3", mode: "query", where: [{ column: "A", op: "in" }] })).toThrow();
    expect(() => parseToolInput("analyze_range" as never, { address: "A1:B3", mode: "query", aggregate: [{ fn: "sum" }] })).toThrow();
    expect(parseToolInput("analyze_range" as never, { address: "A1:B3", mode: "profile" })).toMatchObject({ headerRow: true, limit: 50, where: [], groupBy: [] });
  });

  it("is a tool another open workbook can be asked to run", () => {
    expect(CONNECTED_READ_TOOLS).toContain("analyze_range");
  });
});

describe("analyze_range in the task pane", () => {
  let book: MockWorkbook;
  beforeEach(() => {
    book = new MockWorkbook();
    book.addSheet("Sales");
    book.fill("Sales", [NAMES, ...rows.slice(0, 5).map(row => row.map(value => value as string | number | null)), ...rows.slice(6).map(row => row as Array<string | number | null>)]);
    installOfficeMock(book);
  });
  const call = (args: unknown) => executeTool("analyze_range", parseToolInput("analyze_range" as never, args));

  it("reads the table in the sheet and answers a grouped question", async () => {
    const result = await call({ sheet: { name: "Sales" }, address: "A1:E7", mode: "query", groupBy: ["Product"], aggregate: [{ fn: "sum", column: "Revenue", as: "revenue" }], orderBy: { by: "revenue" } });
    expect(result).toMatchObject({ ok: true, data: { sheet: "Sales", address: "A1:E7", dataRows: 6, truncated: false, matched: 6 } });
    const groups = (result as any).data.groups;
    expect(groups[0]).toMatchObject({ Product: "Pen", revenue: 170 });
    expect(groups.map((group: any) => group.Product)).toEqual(["Pen", "Ink", "Pad"]);
  });

  it("works without a header row, naming columns by their letters", async () => {
    const result = await call({ sheet: { name: "Sales" }, address: "A2:C4", mode: "query", headerRow: false, where: [{ column: "C", op: "gt", value: 6 }], aggregate: [{ fn: "count" }] });
    expect((result as any).data.totals).toEqual({ count: 2 });
  });

  it("says so when the range has nothing below the header, or a column does not exist", async () => {
    expect(await call({ sheet: { name: "Sales" }, address: "A1:E1", mode: "profile" })).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("no data rows") } });
    expect(await call({ sheet: { name: "Sales" }, address: "A1:E7", mode: "query", where: [{ column: "Profit", op: "eq", value: 1 }] })).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("Columns: Region (A)") } });
  });

  it("profiles a sheet in blocks, so a table larger than one read is still covered", async () => {
    const big = book.addSheet("Big");
    const data = Array.from({ length: 6000 }, (_, i) => [`k${i % 7}`, i, i * 2, 1, 0]);
    book.fill("Big", [NAMES, ...data]);
    const result = await call({ sheet: { name: "Big" }, address: "A1:E6001", mode: "query", aggregate: [{ fn: "count" }, { fn: "sum", column: "Units" }] });
    expect(big).toBeDefined();
    expect((result as any).data).toMatchObject({ dataRows: 6000, truncated: false, matched: 6000, totals: { count: 6000, sum_Units: 6000 * 5999 } });
  });
});
