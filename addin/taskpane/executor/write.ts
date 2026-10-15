import {
  EXECUTOR_CAPS,
  numberToColumn,
  parseCell,
  type CellValue,
  type ClearRangeInput,
  type SetFormulaInput,
  type WriteRangeInput,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, stripSheet } from "./common.js";

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

    return {
      sheet: sheet.name,
      address: stripSheet(target.address),
      rowsWritten: rows,
      columnsWritten: columns,
      ...(notes.length > 0 ? { notesWritten: notes.length } : {}),
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
    return { sheet: sheet.name, address: stripSheet(range.address), formula };
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

