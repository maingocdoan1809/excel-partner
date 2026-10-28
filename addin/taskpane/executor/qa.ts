import type { QaApplyResultsInput } from "../../../protocol/src/index.js";
import { ToolFailure } from "./common.js";

/** Output columns this tool may fill, in the order they are appended when missing. */
const OUTPUT_KEYS = ["status", "scriptPath", "lastRun", "lastResult", "duration", "evidence"] as const;

const norm = (value: unknown) => String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");

/** A result must never turn into a formula: text that starts like one is stored as text. */
const safe = (value: string | number | null): string | number => {
  if (value === null) return "";
  return typeof value === "string" && /^[=+\-@]/.test(value) ? `'${value}` : value;
};

/**
 * Write QA run results into the test-case sheet. Rows are matched by the TC_ID column and columns by header
 * text, so the user's own layout is respected. Missing output columns are appended after the last header;
 * every other cell, including formulas in the columns being filled, is left as it was.
 */
export async function applyQaResults(input: QaApplyResultsInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(input.sheet);
    await context.sync();
    if (sheet.isNullObject) throw new ToolFailure("sheet_not_found", `Sheet '${input.sheet}' was not found. Create it from the TestPlan template or pass the sheet name.`);
    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowIndex,columnIndex,rowCount,columnCount,values,formulas");
    await context.sync();
    if (used.isNullObject) throw new ToolFailure("range_not_found", `Sheet '${input.sheet}' is empty. Add a header row with ${input.columns.tcId} first.`);

    const idHeader = norm(input.columns.tcId);
    let headerRow = -1;
    for (let r = 0; r < Math.min(10, used.rowCount) && headerRow < 0; r++) {
      if (used.values[r]!.some((cell) => norm(cell) === idHeader)) headerRow = r;
    }
    if (headerRow < 0) throw new ToolFailure("range_not_found", `No '${input.columns.tcId}' column found in the first rows of '${input.sheet}'.`);

    const columnOf = new Map<string, number>();
    let lastHeader = -1;
    used.values[headerRow]!.forEach((cell, c) => { if (norm(cell) !== "") { columnOf.set(norm(cell), c); lastHeader = c; } });
    const idColumn = columnOf.get(idHeader)!;

    const dataStart = headerRow + 1;
    const dataRows = used.rowCount - dataStart;
    const rowByTc = new Map<string, number>();
    for (let r = dataStart; r < used.rowCount; r++) {
      const key = String(used.values[r]![idColumn] ?? "").trim().toUpperCase();
      if (key && !rowByTc.has(key)) rowByTc.set(key, r);
    }

    const found: string[] = [], notInSheet: string[] = [];
    for (const row of input.rows) (rowByTc.has(row.tcId.toUpperCase()) ? found : notInSheet).push(row.tcId);

    const addedColumns: string[] = [];
    let nextColumn = Math.max(lastHeader, used.columnCount - 1) + 1;
    let updatedCells = 0;
    const headerWrites: Array<{ column: number; text: string }> = [];
    const columnWrites: Array<{ column: number; cells: unknown[][] }> = [];

    for (const key of OUTPUT_KEYS) {
      const header = input.columns[key];
      if (!header || !input.rows.some((row) => row.values[key] !== undefined)) continue;
      let column = columnOf.get(norm(header));
      const existing = column !== undefined;
      if (column === undefined) {
        column = nextColumn++;
        columnOf.set(norm(header), column);
        headerWrites.push({ column, text: header });
        addedColumns.push(header);
      }
      // Start from what the column holds now (formulas included) so only matched rows change.
      const cells: unknown[][] = [];
      for (let r = dataStart; r < used.rowCount; r++) cells.push([existing && column < used.columnCount ? used.formulas[r]![column] : ""]);
      for (const row of input.rows) {
        const index = rowByTc.get(row.tcId.toUpperCase());
        const value = row.values[key];
        if (index === undefined || value === undefined) continue;
        cells[index - dataStart]![0] = safe(value);
        updatedCells++;
      }
      columnWrites.push({ column, cells });
    }

    for (const write of headerWrites) {
      const cell = sheet.getRangeByIndexes(used.rowIndex + headerRow, used.columnIndex + write.column, 1, 1);
      cell.values = [[write.text]];
    }
    if (dataRows > 0) {
      for (const write of columnWrites) {
        sheet.getRangeByIndexes(used.rowIndex + dataStart, used.columnIndex + write.column, dataRows, 1).formulas = write.cells as never;
      }
    }
    await context.sync();
    return {
      sheet: input.sheet, headerRow: used.rowIndex + headerRow + 1, rowsUpdated: found.length, cellsWritten: updatedCells,
      addedColumns, notInSheet, note: notInSheet.length ? `${notInSheet.length} script(s) have no row in the sheet: ${notInSheet.slice(0, 8).join(", ")}${notInSheet.length > 8 ? ", ..." : ""}. Add the test case there with the same ${input.columns.tcId}.` : undefined,
    };
  });
}
