import {
  EXECUTOR_CAPS,
  numberToColumn,
  type CellValue,
  type CopyToInput,
  type GetCellRangesInput,
  type ReadRangeImageInput,
  type ResizeRangeInput,
  type SetCellsInput,
  type Size,
} from "../../../protocol/src/index.js";
import { ToolFailure, normalizeValues, resolveSheet, stripSheet } from "./common.js";
import {
  applyBorders,
  applyCellStyle,
  cellPosition,
  groupByRectangles,
  modalBaseline,
  toBorders,
  toStyle,
  type BorderOut,
  type CellProps,
  type StyleOut,
} from "./styles.js";
import { findOccupied, loadGeometry, overwriteFailure, type Geometry } from "./write.js";

/* -------------------------------------------------------------------------- */
/* get_cell_ranges                                                             */
/* -------------------------------------------------------------------------- */

type CellEntry = CellValue | [CellValue, string] | [CellValue, string, string];

const a1 = (row: number, col: number) => `${numberToColumn(col + 1)}${row + 1}`;

export async function getCellRanges(input: GetCellRangesInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("isNullObject,address");
    await context.sync();

    const notes = used.isNullObject ? new Map<string, string>() : await loadNotes(context, sheet);
    const cells: Record<string, CellEntry> = {};
    /** Style data of the rows actually returned, per chunk; turned into grouped styles at the end. */
    const styled: Array<{ row: number; col: number; props: CellProps[][]; numberFormats: string[][] }> = [];
    let counted = 0;
    let nextRanges: string[] | undefined;

    if (!used.isNullObject) {
      for (let i = 0; i < input.ranges.length; i++) {
        const requested = sheet.getRange(input.ranges[i]!);
        const hit = requested.getIntersectionOrNullObject(used);
        hit.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
        await context.sync();
        if (hit.isNullObject) continue;

        const rowsPerChunk = Math.max(1, Math.floor(EXECUTOR_CAPS.readCells / hit.columnCount));
        for (let offset = 0; offset < hit.rowCount; offset += rowsPerChunk) {
          const rows = Math.min(rowsPerChunk, hit.rowCount - offset);
          const firstRow = hit.rowIndex + offset;
          const chunk = sheet.getRangeByIndexes(firstRow, hit.columnIndex, rows, hit.columnCount);
          chunk.load("values,formulas,numberFormat");
          const props = input.includeStyles ? chunk.getCellProperties(cellPropertiesQuery()) : null;
          await context.sync();

          const values = normalizeValues(chunk.values);
          const keep: number[] = [];
          for (let r = 0; r < rows; r++) {
            let rowCells = 0;
            for (let c = 0; c < hit.columnCount; c++) {
              const formula = String(chunk.formulas[r]?.[c] ?? "");
              const value = values[r]?.[c] ?? null;
              if (formula === "" && (value === null || value === "")) continue;
              const address = a1(firstRow + r, hit.columnIndex + c);
              const note = notes.get(address);
              const entry: CellEntry = formula.startsWith("=")
                ? note !== undefined ? [value, formula, note] : [value, formula]
                : note !== undefined ? [value, "", note] : (value as CellValue);
              cells[address] = entry;
              rowCells++;
            }
            counted += rowCells;
            keep.push(r);
            if (counted >= input.cellLimit) break;
          }

          if (props) {
            styled.push({
              row: firstRow,
              col: hit.columnIndex,
              props: (props.value as CellProps[][]).slice(0, keep.length),
              numberFormats: (chunk.numberFormat as string[][]).slice(0, keep.length),
            });
          }

          if (counted >= input.cellLimit) {
            const remaining: string[] = [];
            const nextRow = firstRow + keep.length;
            const lastRow = hit.rowIndex + hit.rowCount;
            if (nextRow < lastRow) remaining.push(`${a1(nextRow, hit.columnIndex)}:${a1(lastRow - 1, hit.columnIndex + hit.columnCount - 1)}`);
            remaining.push(...input.ranges.slice(i + 1));
            if (remaining.length > 0) nextRanges = remaining;
            i = input.ranges.length; // leave the outer loop
            break;
          }
        }
      }
    }

    const result: Record<string, unknown> = {
      sheet: sheet.name,
      dimension: used.isNullObject ? null : stripSheet(used.address),
      cells,
      hasMore: nextRanges !== undefined,
      ...(nextRanges ? { nextRanges } : {}),
    };

    if (input.includeStyles) {
      const base = modalBaseline(styled.flatMap((chunk) => chunk.props));
      const styles: Record<string, StyleOut> = {};
      const borders: Record<string, BorderOut> = {};
      for (const { row, col, props, numberFormats } of styled) {
        const styleGrid = props.map((cells, r) => cells.map((p, c) => toStyle(p, numberFormats[r]?.[c], base)));
        Object.assign(styles, groupByRectangles(styleGrid, row, col));
        Object.assign(borders, groupByRectangles(props.map((cells) => cells.map(toBorders)), row, col));
      }
      result.styles = styles;
      result.borders = borders;
      result.baseline = base;
    }
    return result;
  });
}

function cellPropertiesQuery(): Excel.CellPropertiesLoadOptions {
  const side = { style: true, color: true, weight: true };
  return {
    format: {
      font: { bold: true, italic: true, color: true, name: true, size: true, underline: true, strikethrough: true },
      fill: { color: true },
      horizontalAlignment: true,
      borders: { top: side, bottom: side, left: side, right: side },
    },
  } as Excel.CellPropertiesLoadOptions;
}

/** Address -> note text for the sheet; empty when this Excel has no notes API or none exist. */
async function loadNotes(context: Excel.RequestContext, sheet: Excel.Worksheet): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.18")) return out;
  try {
    sheet.notes.load("items/content");
    await context.sync();
    const located = sheet.notes.items.slice(0, 2_000).map((note) => {
      const where = note.getLocation();
      where.load("address");
      return { note, where };
    });
    if (located.length > 0) await context.sync();
    for (const { note, where } of located) out.set(stripSheet(where.address), note.content);
  } catch {
    /* notes are optional context: never fail a read because of them */
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* set_cells                                                                   */
/* -------------------------------------------------------------------------- */

export async function setCells(input: SetCellsInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const entries = Object.entries(input.cells).map(([address, spec]) => ({ address: address.replace(/\$/g, ""), spec, pos: cellPosition(address) }));

    // Everything that can be refused is checked before the first change is queued.
    if (entries.some((e) => e.spec.note !== undefined) && !Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
      throw new ToolFailure("unsupported", "this Excel version cannot write cell notes (needs ExcelApi 1.18)");
    }

    const bounds = entries.reduce(
      (b, e) => ({
        r0: Math.min(b.r0, e.pos.row),
        c0: Math.min(b.c0, e.pos.col),
        r1: Math.max(b.r1, e.pos.row),
        c1: Math.max(b.c1, e.pos.col),
      }),
      { r0: Infinity, c0: Infinity, r1: -1, c1: -1 },
    );
    const source = sheet.getRangeByIndexes(bounds.r0, bounds.c0, bounds.r1 - bounds.r0 + 1, bounds.c1 - bounds.c0 + 1);

    if (!input.allowOverwrite) {
      const writing = entries.filter((e) => e.spec.value !== undefined || e.spec.formula !== undefined);
      const targets = writing.map((e) => ({
        range: sheet.getRange(e.address),
        geometry: { rowIndex: e.pos.row, columnIndex: e.pos.col, rowCount: 1, columnCount: 1 } as Geometry,
      }));
      let destination: { range: Excel.Range; geometry: Geometry } | null = null;
      if (input.copyTo) {
        const range = sheet.getRange(input.copyTo);
        destination = { range, geometry: await loadGeometry(context, range) };
        targets.push(destination);
      }
      const inSource = (r: number, c: number) => r >= bounds.r0 && r <= bounds.r1 && c >= bounds.c0 && c <= bounds.c1;
      // Source cells are rewritten by this very call, so the destination check skips the source block
      // while the per-cell check covers whatever is currently in those cells.
      const found = await findOccupied(
        context,
        targets.filter((t) => t !== destination),
      );
      if (destination) {
        const fromDestination = await findOccupied(context, [destination], inSource);
        found.total += fromDestination.total;
        found.cells.push(...fromDestination.cells.slice(0, Math.max(0, EXECUTOR_CAPS.overwritePreview - found.cells.length)));
      }
      if (found.total > 0) throw overwriteFailure(found);
    }

    const formulaCells: string[] = [];
    const noteEntries: Array<[string, string]> = [];
    for (const { address, spec } of entries) {
      const cell = sheet.getRange(address);
      if (spec.formula !== undefined) {
        cell.formulas = [[spec.formula.startsWith("=") ? spec.formula : `=${spec.formula}`]];
        formulaCells.push(address);
      } else if (spec.value !== undefined) {
        cell.values = [[spec.value]];
      }
      if (spec.style) {
        applyCellStyle(cell, spec.style);
        if (spec.style.numberFormat) cell.numberFormat = [[spec.style.numberFormat]];
      }
      if (spec.borders) applyBorders(cell, spec.borders);
      if (spec.note !== undefined) noteEntries.push([address, spec.note]);
    }
    if (noteEntries.length > 0) await writeNotesFor(context, sheet, noteEntries);

    let filled: Excel.Range = source;
    if (input.copyTo) {
      filled = sheet.getRange(input.copyTo);
      filled.copyFrom(source, Excel.RangeCopyType.all);
    }

    const messages: string[] = [];
    if (input.resizeWidth) applySize(source.getEntireColumn(), input.resizeWidth, "width");
    if (input.resizeHeight) applySize(source.getEntireRow(), input.resizeHeight, "height");
    await context.sync();

    // Report what the formulas evaluated to, so the model can catch #REF!/#NAME? straight away.
    const formulaResults: Record<string, number | string> = {};
    const probes = formulaCells.slice(0, 200).map((address) => ({ address, range: sheet.getRange(address) }));
    probes.forEach((p) => p.range.load("values"));
    let copied: Excel.Range | null = null;
    if (input.copyTo) {
      copied = filled;
      copied.load("values,formulas,rowIndex,columnIndex");
    }
    await context.sync();
    for (const { address, range } of probes) {
      const value = range.values[0]?.[0];
      if (value !== null && value !== undefined && value !== "") formulaResults[address] = value as number | string;
    }
    if (copied) {
      let shown = 0;
      copied.formulas.forEach((row, r) =>
        row.forEach((formula, c) => {
          if (shown >= 200 || !String(formula).startsWith("=")) return;
          const value = copied!.values[r]?.[c];
          const address = a1(copied!.rowIndex + r, copied!.columnIndex + c);
          if (!(address in formulaResults) && value !== null && value !== "") {
            formulaResults[address] = value as number | string;
            shown++;
          }
        }),
      );
    }
    const errors = Object.entries(formulaResults).filter(([, v]) => typeof v === "string" && v.startsWith("#"));
    if (errors.length > 0) messages.push(`${errors.length} formula(s) evaluated to an error: ${errors.slice(0, 5).map(([k, v]) => `${k}=${v}`).join(", ")}`);

    return {
      sheet: sheet.name,
      cellsWritten: entries.length,
      ...(input.copyTo ? { copiedTo: input.copyTo } : {}),
      formulaResults,
      ...(messages.length > 0 ? { messages } : {}),
    };
  });
}

async function writeNotesFor(context: Excel.RequestContext, sheet: Excel.Worksheet, notes: Array<[string, string]>): Promise<void> {
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

/* -------------------------------------------------------------------------- */
/* copy_to, resize_range, read_range_image                                     */
/* -------------------------------------------------------------------------- */

export async function copyTo(input: CopyToInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const source = sheet.getRange(input.source);
    const destination = sheet.getRange(input.destination);
    const src = await loadGeometry(context, source);
    const dest = await loadGeometry(context, destination);

    // Formats-only copies never touch cell contents.
    if (input.copyType !== "formats" && !input.allowOverwrite) {
      const inSource = (r: number, c: number) =>
        r >= src.rowIndex && r < src.rowIndex + src.rowCount && c >= src.columnIndex && c < src.columnIndex + src.columnCount;
      const found = await findOccupied(context, [{ range: destination, geometry: dest }], inSource);
      if (found.total > 0) throw overwriteFailure(found);
    }

    const type = { all: Excel.RangeCopyType.all, formulas: Excel.RangeCopyType.formulas, values: Excel.RangeCopyType.values, formats: Excel.RangeCopyType.formats }[input.copyType];
    destination.copyFrom(source, type);
    destination.load("address");
    await context.sync();
    return { sheet: sheet.name, source: input.source, address: stripSheet(destination.address), copyType: input.copyType };
  });
}

function applySize(range: Excel.Range, size: Size, axis: "width" | "height"): void {
  const fmt = range.format;
  if (axis === "width") {
    if (size.type === "points") fmt.columnWidth = size.value!;
    else if (size.type === "standard") fmt.useStandardWidth = true;
    else fmt.autofitColumns();
  } else if (size.type === "points") fmt.rowHeight = size.value!;
  else if (size.type === "standard") fmt.useStandardHeight = true;
  else fmt.autofitRows();
}

export async function resizeRange(input: ResizeRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = input.address ? sheet.getRange(input.address) : sheet.getRange();
    if (input.width) applySize(range.getEntireColumn(), input.width, "width");
    if (input.height) applySize(range.getEntireRow(), input.height, "height");
    range.load("address");
    await context.sync();
    return { sheet: sheet.name, address: input.address ? stripSheet(range.address) : "entire sheet", width: input.width ?? null, height: input.height ?? null };
  });
}

/** Largest image (base64 chars) returned to the model; bigger renders are refused with advice. */
const MAX_IMAGE_BASE64 = 1_500_000;

export async function readRangeImage(input: ReadRangeImageInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    range.load("address");
    const image = range.getImage();
    await context.sync();
    if (image.value.length > MAX_IMAGE_BASE64) {
      throw new ToolFailure("size_limit_exceeded", `rendered image of ${input.address} is too large; render a smaller range (about 50 rows or fewer)`);
    }
    return { sheet: sheet.name, address: stripSheet(range.address), image: { mimeType: "image/png", data: image.value } };
  });
}
