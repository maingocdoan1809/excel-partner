import type { FormatRangeInput } from "../../../protocol/src/index.js";
import { resolveSheet, stripSheet, runExcel } from "./common.js";
import { MAX_COLUMN, MAX_ROW } from "../../../protocol/src/index.js";
import { overRows } from "./bulk.js";

/** Measured in real Excel: formatting about a million cells takes ~0.6 s in one round trip, so only bigger blocks are sliced. */
const FORMAT_SLICE_CELLS = 500_000;

/** A whole column or row: Excel formats these in one cheap operation, so they are never sliced. */
const isEntireLine = (rows: number, columns: number): boolean => rows >= MAX_ROW || columns >= MAX_COLUMN;

/** The settings applied to every cell, in the order they are reported. */
const CELL_SETTINGS = ["bold", "italic", "underline", "fontSize", "fontColor", "fillColor", "numberFormat", "horizontalAlignment", "verticalAlignment", "wrapText"] as const;

const OUTER_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;

export async function formatRange(input: FormatRangeInput): Promise<unknown> {
  return runExcel(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    range.load("address,rowCount,columnCount");
    await context.sync();

    const f = input.format;

    /** Settings that belong to each cell: the same for any slice of rows. */
    const applyCellFormat = (target: Excel.Range) => {
      const fmt = target.format;
      if (f.bold !== undefined) fmt.font.bold = f.bold;
      if (f.italic !== undefined) fmt.font.italic = f.italic;
      if (f.underline !== undefined) fmt.font.underline = f.underline ? Excel.RangeUnderlineStyle.single : Excel.RangeUnderlineStyle.none;
      if (f.fontSize !== undefined) fmt.font.size = f.fontSize;
      if (f.fontColor !== undefined) fmt.font.color = f.fontColor;
      if (f.fillColor !== undefined) fmt.fill.color = f.fillColor;
      // One value is applied to every cell of the range, so no grid of strings has to be built and sent.
      if (f.numberFormat !== undefined) target.numberFormat = f.numberFormat as unknown as string[][];
      if (f.horizontalAlignment !== undefined) fmt.horizontalAlignment = f.horizontalAlignment;
      if (f.verticalAlignment !== undefined) fmt.verticalAlignment = f.verticalAlignment;
      if (f.wrapText !== undefined) fmt.wrapText = f.wrapText;
    };
    const applied: string[] = CELL_SETTINGS.filter((name) => f[name] !== undefined);

    // A big block is formatted slice by slice so Excel stays responsive. Formatting never changes a value, so
    // calculation is left alone.
    if (range.rowCount * range.columnCount > FORMAT_SLICE_CELLS && !isEntireLine(range.rowCount, range.columnCount)) {
      await overRows(context, range, range.rowCount, range.columnCount, "fill", (slice) => applyCellFormat(slice));
    } else {
      applyCellFormat(range);
    }

    // Whole-range settings (a border around the block, the width of a column) are applied once, not per slice.
    const fmt = range.format;
    if (f.columnWidth !== undefined) { fmt.columnWidth = f.columnWidth; applied.push("columnWidth"); }
    if (f.rowHeight !== undefined) { fmt.rowHeight = f.rowHeight; applied.push("rowHeight"); }
    if (f.borders !== undefined) { applyBorders(range, f.borders); applied.push("borders"); }
    if (f.autofitColumns) { fmt.autofitColumns(); applied.push("autofitColumns"); }

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
