import {
  EXECUTOR_CAPS,
  numberToColumn,
  parseCell,
  type CellValue,
  type ClearRangeInput,
  type SetFormulaInput,
  type WriteRangeInput,
} from "../../../protocol/src/index.js";
import { ToolFailure, isErrorValue, resolveSheet, stripSheet } from "./common.js";

/** Cells re-read after a write to report what formulas evaluated to. */
const VERIFY_CELL_LIMIT = 2_000;
/** Entries returned in `formulaResults`; errors beyond this are still counted. */
const VERIFY_RESULT_LIMIT = 200;

export interface FormulaReport {
  formulaResults?: Record<string, number | string>;
  messages?: string[];
}

/**
 * Re-read a just-written block and report what its formulas evaluated to, so the model can catch
 * `#REF!` / `#NAME?` without another call. Never fails the write: a failed read-back is reported
 * as a message instead.
 */
export async function reportFormulaResults(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  geometry: Geometry,
  extraMessages: string[] = [],
): Promise<FormulaReport> {
  const messages = [...extraMessages];
  const formulaResults: Record<string, number | string> = {};
  try {
    const columns = Math.max(1, geometry.columnCount);
    const rows = Math.min(geometry.rowCount, Math.max(1, Math.floor(VERIFY_CELL_LIMIT / columns)));
    const block = sheet.getRangeByIndexes(geometry.rowIndex, geometry.columnIndex, rows, geometry.columnCount);
    block.load("formulas,values");
    await context.sync();

    const errors: string[] = [];
    let listed = 0;
    block.formulas.forEach((row, r) =>
      row.forEach((formula, c) => {
        if (!String(formula).startsWith("=")) return;
        const value = block.values[r]?.[c];
        const address = `${numberToColumn(geometry.columnIndex + c + 1)}${geometry.rowIndex + r + 1}`;
        if (isErrorValue(value)) errors.push(`${address}=${value}`);
        if (listed < VERIFY_RESULT_LIMIT && value !== null && value !== undefined && value !== "") {
          formulaResults[address] = value as number | string;
          listed++;
        }
      }),
    );
    if (errors.length > 0) {
      messages.push(`${errors.length} formula(s) evaluated to an error: ${errors.slice(0, 5).join(", ")}${errors.length > 5 ? ", ..." : ""}. Fix them before reporting success.`);
    }
    if (rows < geometry.rowCount) {
      messages.push(`Only the first ${rows} of ${geometry.rowCount} rows were checked for formula errors; use find_errors for the rest.`);
    }
  } catch (error) {
    messages.push(`Could not read the formula results back (${error instanceof Error ? error.message : String(error)}); verify with read_range.`);
  }
  return {
    ...(Object.keys(formulaResults).length > 0 ? { formulaResults } : {}),
    ...(messages.length > 0 ? { messages } : {}),
  };
}

/** Zero-based position and size of a range, as loaded from Office.js. */
export interface Geometry {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

export interface Occupancy {
  total: number;
  /** First few non-empty cell addresses, for the error message. */
  cells: string[];
}

/**
 * Find non-empty cells (constants or any formula, even one evaluating to ""; formatting alone
 * does not count) across several ranges with one sync. `skip` excludes cells that are about to
 * be written anyway (e.g. the source of a copy). Reading failures propagate: the callers must
 * fail closed, because a swallowed error here would silently skip the overwrite guard.
 */
export async function findOccupied(
  context: Excel.RequestContext,
  targets: Array<{ range: Excel.Range; geometry: Geometry }>,
  skip?: (row: number, col: number) => boolean,
): Promise<Occupancy> {
  targets.forEach(({ range }) => range.load("formulas"));
  await context.sync();

  const found: Occupancy = { total: 0, cells: [] };
  for (const { range, geometry } of targets) {
    range.formulas.forEach((row, r) =>
      row.forEach((cell, c) => {
        if (cell === "" || cell === null || cell === undefined) return;
        const absRow = geometry.rowIndex + r;
        const absCol = geometry.columnIndex + c;
        if (skip?.(absRow, absCol)) return;
        found.total++;
        if (found.cells.length < EXECUTOR_CAPS.overwritePreview) found.cells.push(`${numberToColumn(absCol + 1)}${absRow + 1}`);
      }),
    );
  }
  return found;
}

export function overwriteFailure(found: Occupancy): ToolFailure {
  const more = found.total > found.cells.length ? `, ... (${found.total - found.cells.length} more)` : "";
  return new ToolFailure(
    "would_overwrite",
    `Would overwrite ${found.total} non-empty cell${found.total === 1 ? "" : "s"}: ${found.cells.join(", ")}${more}. ` +
      "Read them with read_range and ask the user to confirm; then retry with allowOverwrite=true.",
    { nonEmpty: found.total, cells: found.cells },
  );
}

/** Refuse to overwrite non-empty cells unless the caller opted in. */
export async function assertCanWrite(
  context: Excel.RequestContext,
  target: Excel.Range,
  geometry: Geometry,
  allowOverwrite: boolean,
  skip?: (row: number, col: number) => boolean,
): Promise<void> {
  if (allowOverwrite) return;
  const found = await findOccupied(context, [{ range: target, geometry }], skip);
  if (found.total > 0) throw overwriteFailure(found);
}

export async function loadGeometry(context: Excel.RequestContext, range: Excel.Range): Promise<Geometry> {
  range.load("rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
  return {
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
  };
}

export async function writeRange(input: WriteRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");

    const rows = input.values.length;
    const columns = input.values.reduce((max, row) => Math.max(max, row.length), 0);
    // Office.js needs a rectangle; `null` means "leave this cell alone".
    const grid: Array<Array<CellValue | null>> = input.values.map((row) => [
      ...row,
      ...Array<null>(columns - row.length).fill(null),
    ]);

    const anchor = sheet.getRange(input.address);
    const anchorGeometry = await loadGeometry(context, anchor);
    const anchorCells = anchorGeometry.rowCount * anchorGeometry.columnCount;
    if (input.mode === "values" && anchorCells > 1 && (anchorGeometry.rowCount !== rows || anchorGeometry.columnCount !== columns)) {
      throw new ToolFailure(
        "invalid_arguments",
        `address ${input.address} is ${anchorGeometry.rowCount}x${anchorGeometry.columnCount} but values are ${rows}x${columns}; ` +
          "use a matching address, a single anchor cell, or mode `top_left`.",
      );
    }

    const target = anchor.getCell(0, 0).getResizedRange(rows - 1, columns - 1);
    const geometry: Geometry = { ...anchorGeometry, rowCount: rows, columnCount: columns };

    // Validate everything that can fail before the first change is queued.
    const notes = Object.entries(input.notes ?? {});
    if (notes.length > 0) {
      if (!Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
        throw new ToolFailure("unsupported", "this Excel version cannot write cell notes (needs ExcelApi 1.18)");
      }
      for (const [address] of notes) assertCellInside(address, geometry);
    }
    await assertCanWrite(context, target, geometry, input.allowOverwrite);

    target.values = grid;
    if (notes.length > 0) await writeNotes(context, sheet, notes);
    target.load("address");
    await context.sync();

    const formulaStrings = grid.flat().filter((cell) => typeof cell === "string" && cell.startsWith("=")).length;
    const report = formulaStrings > 0
      ? await reportFormulaResults(context, sheet, geometry, [
          `${formulaStrings} value(s) starting with "=" were stored as formulas. To store literal text, prefix it with an apostrophe.`,
        ])
      : {};

    return {
      sheet: sheet.name,
      address: stripSheet(target.address),
      rowsWritten: rows,
      columnsWritten: columns,
      ...(notes.length > 0 ? { notesWritten: notes.length } : {}),
      ...report,
    };
  });
}

function assertCellInside(address: string, geometry: Geometry): void {
  const cell = parseCell(address);
  const inside =
    cell !== null &&
    cell.row - 1 >= geometry.rowIndex &&
    cell.row - 1 < geometry.rowIndex + geometry.rowCount &&
    cell.col - 1 >= geometry.columnIndex &&
    cell.col - 1 < geometry.columnIndex + geometry.columnCount;
  if (!inside) {
    throw new ToolFailure("invalid_arguments", `note address '${address}' is not a single cell inside the written range`);
  }
}

async function writeNotes(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  notes: Array<[string, string]>,
): Promise<void> {
  const existing = notes.map(([address]) => {
    const note = sheet.notes.getItemOrNullObject(address);
    note.load("isNullObject");
    return note;
  });
  await context.sync();
  notes.forEach(([address, text], i) => {
    const note = existing[i]!;
    if (note.isNullObject) sheet.notes.add(address, text);
    else note.content = text;
  });
}

export async function setFormula(input: SetFormulaInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    const geometry = await loadGeometry(context, range);
    await assertCanWrite(context, range, geometry, input.allowOverwrite);

    const formula = input.formula.startsWith("=") ? input.formula : `=${input.formula}`;
    const first = range.getCell(0, 0);
    first.formulas = [[formula]];
    // Fill from the first cell so relative references shift per cell.
    if (geometry.rowCount * geometry.columnCount > 1) {
      range.copyFrom(first, Excel.RangeCopyType.formulas);
    }
    range.load("address");
    await context.sync();
    const report = await reportFormulaResults(context, sheet, geometry);
    return { sheet: sheet.name, address: stripSheet(range.address), formula, ...report };
  });
}

export async function clearRange(input: ClearRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    range.load("address");
    const applyTo = {
      contents: Excel.ClearApplyTo.contents,
      formats: Excel.ClearApplyTo.formats,
      all: Excel.ClearApplyTo.all,
    }[input.clearType];
    range.clear(applyTo);
    await context.sync();
    return { sheet: sheet.name, address: stripSheet(range.address), clearType: input.clearType };
  });
}

