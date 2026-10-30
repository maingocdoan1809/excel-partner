import { numberToColumn, type RecordTestResultsInput } from "../../../protocol/src/index.js";
import { ToolFailure } from "./common.js";

const norm = (value: unknown) => String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");

/** Results come from the system under test, so text that looks like a formula is stored as text. */
const safe = (value: string | number | null): string | number => {
  if (value === null) return "";
  return typeof value === "string" && /^[=+\-@]/.test(value) ? `'${value}` : value;
};

/** What this tool wrote in this task pane, so later updates of its own cells (Đang test -> Pass) need no question. */
const written = new Map<string, string>();
const cellKey = (sheet: string, row: number, column: number) => `${sheet}|${row}|${column}`;

/**
 * Write progress and results into a test-case sheet of any layout. Rows are given by number and columns by header
 * text, so nothing about the sheet is fixed in advance. Missing columns are appended after the last header; a cell that
 * already holds something this tool did not write is never touched unless the caller says the user agreed.
 */
export async function recordTestResults(input: RecordTestResultsInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(input.sheet);
    await context.sync();
    if (sheet.isNullObject) throw new ToolFailure("sheet_not_found", `Sheet '${input.sheet}' was not found.`);
    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowIndex,rowCount,columnIndex,columnCount");
    await context.sync();
    if (used.isNullObject) throw new ToolFailure("range_not_found", `Sheet '${input.sheet}' is empty.`);
    const lastRow = used.rowIndex + used.rowCount;
    const lastColumn = used.columnIndex + used.columnCount;
    if (input.headerRow > lastRow) throw new ToolFailure("invalid_arguments", `headerRow ${input.headerRow} is below the last used row (${lastRow}).`);

    const headerRange = sheet.getRangeByIndexes(input.headerRow - 1, 0, 1, lastColumn);
    headerRange.load("values");
    await context.sync();
    const headers = headerRange.values[0] ?? [];
    const columnOf = new Map<string, number>();
    let lastHeader = -1;
    headers.forEach((cell, index) => { if (norm(cell) !== "") { if (!columnOf.has(norm(cell))) columnOf.set(norm(cell), index); lastHeader = index; } });
    if (lastHeader < 0) throw new ToolFailure("range_not_found", `Row ${input.headerRow} of '${input.sheet}' has no header text. Read the sheet and pass the row that holds the column names.`);

    for (const item of input.results) {
      if (item.row <= input.headerRow || item.row > lastRow) {
        throw new ToolFailure("invalid_arguments", `Row ${item.row} is not a data row (data rows are ${input.headerRow + 1}..${lastRow}).`);
      }
    }

    const wanted = new Map<string, string>();
    for (const item of input.results) for (const key of Object.keys(item.values)) if (!wanted.has(norm(key))) wanted.set(norm(key), key.trim());
    const missing = [...wanted].filter(([key]) => !columnOf.has(key));
    if (missing.length && !input.addMissingColumns) {
      throw new ToolFailure("invalid_arguments", `No column named ${missing.map(([, text]) => `'${text}'`).join(", ")} in row ${input.headerRow}. Existing: ${[...columnOf.keys()].slice(0, 30).join(", ")}.`);
    }
    const addedColumns: string[] = [];
    let nextColumn = lastHeader + 1;
    const newColumns: Array<{ column: number; text: string }> = [];
    for (const [key, text] of missing) {
      columnOf.set(key, nextColumn);
      newColumns.push({ column: nextColumn, text });
      addedColumns.push(`${text} (${numberToColumn(nextColumn + 1)})`);
      nextColumn++;
    }

    const targets: Array<{ row: number; column: number; value: string | number | null; range: Excel.Range }> = [];
    for (const item of input.results) {
      for (const [key, value] of Object.entries(item.values)) {
        const column = columnOf.get(norm(key))!;
        targets.push({ row: item.row, column, value, range: sheet.getRangeByIndexes(item.row - 1, column, 1, 1) });
      }
    }
    targets.forEach(({ range }) => range.load("formulas,values"));
    await context.sync();

    if (!input.allowOverwrite) {
      const occupied = targets.filter(({ row, column, range }) => {
        const formula = range.formulas[0]![0];
        if (formula === "" || formula === null || formula === undefined) return false;
        return written.get(cellKey(input.sheet, row, column)) !== String(range.values[0]![0]);
      });
      if (occupied.length) {
        const cells = occupied.slice(0, 10).map(({ row, column }) => `${numberToColumn(column + 1)}${row}`);
        throw new ToolFailure("would_overwrite",
          `Would overwrite ${occupied.length} non-empty cell${occupied.length === 1 ? "" : "s"}: ${cells.join(", ")}${occupied.length > cells.length ? ", ..." : ""}. `
          + "Read them with read_range and ask the user; only if they agree, retry with allowOverwrite=true.",
          { nonEmpty: occupied.length, cells });
      }
    }

    if (newColumns.length) {
      const model = sheet.getRangeByIndexes(input.headerRow - 1, lastHeader, 1, 1);
      // Only the look of the neighbouring header is copied, not its merge or borders, so a merged title cell is never spread.
      const look = model.getCellProperties({
        format: { font: { bold: true, color: true, name: true, size: true }, fill: { color: true }, horizontalAlignment: true, verticalAlignment: true, wrapText: true },
      });
      await context.sync();
      const format = look.value[0]![0]!.format!;
      for (const { column, text } of newColumns) {
        const cell = sheet.getRangeByIndexes(input.headerRow - 1, column, 1, 1);
        cell.values = [[text]];
        if (format.fill?.color && format.fill.color.toUpperCase() !== "#FFFFFF") cell.format.fill.color = format.fill.color;
        if (format.font?.bold !== undefined) cell.format.font.bold = format.font.bold;
        if (format.font?.color) cell.format.font.color = format.font.color;
        if (format.font?.name) cell.format.font.name = format.font.name;
        if (format.font?.size) cell.format.font.size = format.font.size;
        if (format.horizontalAlignment) cell.format.horizontalAlignment = format.horizontalAlignment;
        if (format.verticalAlignment) cell.format.verticalAlignment = format.verticalAlignment;
        if (format.wrapText !== undefined) cell.format.wrapText = format.wrapText;
      }
    }
    for (const target of targets) target.range.values = [[safe(target.value)]];
    await context.sync();
    for (const { row, column, value } of targets) written.set(cellKey(input.sheet, row, column), String(value ?? ""));

    return {
      sheet: input.sheet, headerRow: input.headerRow, rowsUpdated: new Set(input.results.map(item => item.row)).size, cellsWritten: targets.length,
      addedColumns, columns: Object.fromEntries([...wanted].map(([key, text]) => [text, numberToColumn(columnOf.get(key)! + 1)])),
    };
  });
}
