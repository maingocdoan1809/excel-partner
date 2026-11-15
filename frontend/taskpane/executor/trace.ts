import { splitSheetAddress, type TraceFormulaInput } from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, runExcel, stripSheet } from "./common.js";

/** A range whose own precedents or dependents are followed in the next step: bigger ones are shown but not opened up. */
const EXPAND_CELLS = 400;

interface TraceNode {
  direction: "precedents" | "dependents";
  level: number;
  address: string;
  cells: number;
  /** Only for a single cell. */
  formula?: string;
  value?: unknown;
}

const key = (address: string): string => address.replace(/\$/g, "").replace(/^'|'(?=!)/g, "").toUpperCase();

/**
 * `trace_formula`: what a cell's formula reads, and what reads the cell, a few steps deep. The two directions are
 * Excel's own (`Range.getDirectPrecedents` / `getDirectDependents`, ExcelApi 1.12), so they follow other sheets too.
 */
export async function traceFormula(input: TraceFormulaInput): Promise<unknown> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
    throw new ToolFailure("unsupported", "this Excel version cannot trace formulas (needs ExcelApi 1.12, Microsoft 365)");
  }
  return runExcel(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const origin = sheet.getRange(input.address);
    origin.load("address,formulas,values");
    await context.sync();

    const at = (address: string): Excel.Range => {
      const { sheet: name, range } = splitSheetAddress(address);
      return (name ? context.workbook.worksheets.getItem(name) : sheet).getRange(range);
    };
    const nodes: TraceNode[] = [];
    const notes: string[] = [];
    let truncated = false;
    const formula = String(origin.formulas[0]?.[0] ?? "");
    if (!formula.startsWith("=") && input.direction !== "dependents") notes.push(`${stripSheet(origin.address)} holds no formula, so it has no precedents.`);

    for (const direction of input.direction === "both" ? ["precedents", "dependents"] as const : [input.direction]) {
      const seen = new Set<string>([key(origin.address)]);
      let frontier: Excel.Range[] = [origin];
      for (let level = 1; level <= input.depth && frontier.length && !truncated; level++) {
        const areas = frontier.map((range) => (direction === "precedents" ? range.getDirectPrecedents() : range.getDirectDependents()));
        areas.forEach((area) => area.load("addresses"));
        await context.sync();
        const found: string[] = [];
        for (const area of areas) for (const address of area.addresses) if (!seen.has(key(address))) { seen.add(key(address)); found.push(address); }
        const room = input.limit - nodes.length;
        if (found.length > room) { truncated = true; found.length = Math.max(0, room); }
        const ranges = found.map(at);
        ranges.forEach((range) => range.load("address,rowCount,columnCount"));
        await context.sync();
        const singles = ranges.filter((range) => range.rowCount * range.columnCount === 1);
        singles.forEach((range) => range.load("formulas,values"));
        await context.sync();
        for (const range of ranges) {
          const cells = range.rowCount * range.columnCount;
          const node: TraceNode = { direction, level, address: range.address, cells };
          if (cells === 1) {
            const text = String(range.formulas[0]?.[0] ?? "");
            if (text.startsWith("=")) node.formula = text;
            node.value = range.values[0]?.[0] ?? null;
          }
          nodes.push(node);
        }
        frontier = ranges.filter((range) => range.rowCount * range.columnCount <= EXPAND_CELLS);
        if (frontier.length < ranges.length) notes.push(`Ranges of more than ${EXPAND_CELLS} cells are listed but not followed further.`);
      }
    }
    if (truncated) notes.push(`Stopped at ${input.limit} cells and ranges; trace a nearer cell or lower the depth for the rest.`);
    return {
      sheet: sheet.name, address: stripSheet(origin.address), ...(formula.startsWith("=") ? { formula } : {}), value: origin.values[0]?.[0] ?? null,
      direction: input.direction, depth: input.depth, nodes, truncated, notes: [...new Set(notes)],
    };
  });
}
