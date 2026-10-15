import type { FormatRangeInput } from "../../../protocol/src/index.js";
import { resolveSheet, stripSheet } from "./common.js";

const OUTER_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;

export async function formatRange(input: FormatRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    range.load("address,rowCount,columnCount");
    await context.sync();

    const fmt = range.format;
    const f = input.format;
    const applied: string[] = [];
    const apply = (name: string, run: () => void) => {
      run();
      applied.push(name);
    };

    if (f.bold !== undefined) apply("bold", () => (fmt.font.bold = f.bold!));
    if (f.italic !== undefined) apply("italic", () => (fmt.font.italic = f.italic!));
    if (f.underline !== undefined) {
      apply("underline", () => {
        fmt.font.underline = f.underline ? Excel.RangeUnderlineStyle.single : Excel.RangeUnderlineStyle.none;
      });
    }
    if (f.fontSize !== undefined) apply("fontSize", () => (fmt.font.size = f.fontSize!));
    if (f.fontColor !== undefined) apply("fontColor", () => (fmt.font.color = f.fontColor!));
    if (f.fillColor !== undefined) apply("fillColor", () => (fmt.fill.color = f.fillColor!));
    if (f.numberFormat !== undefined) {
      // Office.js wants one entry per cell, not a single value.
      const grid = Array.from({ length: range.rowCount }, () => Array<string>(range.columnCount).fill(f.numberFormat!));
      apply("numberFormat", () => (range.numberFormat = grid));
    }
    if (f.horizontalAlignment !== undefined) {
      apply("horizontalAlignment", () => (fmt.horizontalAlignment = f.horizontalAlignment!));
    }
    if (f.verticalAlignment !== undefined) {
      apply("verticalAlignment", () => (fmt.verticalAlignment = f.verticalAlignment!));
    }
    if (f.wrapText !== undefined) apply("wrapText", () => (fmt.wrapText = f.wrapText!));
    if (f.columnWidth !== undefined) apply("columnWidth", () => (fmt.columnWidth = f.columnWidth!));
    if (f.rowHeight !== undefined) apply("rowHeight", () => (fmt.rowHeight = f.rowHeight!));
    if (f.borders !== undefined) apply("borders", () => applyBorders(range, f.borders!));
    if (f.autofitColumns) apply("autofitColumns", () => fmt.autofitColumns());

    await context.sync();
    return { sheet: sheet.name, address: stripSheet(range.address), applied };
  });
}

function applyBorders(range: Excel.Range, mode: "all" | "outline" | "none"): void {
  const style = mode === "none" ? Excel.BorderLineStyle.none : Excel.BorderLineStyle.continuous;
  const edges: string[] = [...OUTER_EDGES];
  if (mode !== "outline") {
    // Inside borders only exist for ranges with more than one row/column.
    if (range.rowCount > 1) edges.push("InsideHorizontal");
    if (range.columnCount > 1) edges.push("InsideVertical");
  }
  for (const edge of edges) {
    const border = range.format.borders.getItem(edge as Excel.BorderIndex);
    border.style = style;
    if (mode !== "none") border.color = "#000000";
  }
}
