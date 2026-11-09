import { resolveSheet } from "./common.js";
import type { LayoutSpec, UndoPlan, UndoTarget } from "./undo-plan.js";

/**
 * The Excel side of undo: photograph the cells a tool is about to change, fingerprint them afterwards, and put
 * them back later. Measured in real Excel: values/formulas/number formats cost ~2 ms per 1,000 cells, styles ~90 ms.
 */

export const PHOTO_LIMITS = { cells: 50_000, styledCells: 20_000, columns: 300, rows: 2_000 } as const;

/** The change is too big to keep a photo of; the tool still runs, but it cannot be undone. */
export class TooLarge extends Error {
  override readonly name = "TooLarge";
}

type Grid<T> = T[][];

export interface Block {
  sheetId: string;
  sheetName: string;
  /** Zero-based top-left corner and size. */
  row: number;
  col: number;
  rows: number;
  cols: number;
  formulas: Grid<unknown>;
  valueTypes: Grid<string>;
  numberFormat: Grid<string>;
  styles?: { palette: string[]; grid: number[][] };
  /** Merged areas as [row, col, rows, cols], zero-based, absolute. */
  merged?: Array<[number, number, number, number]>;
}

export interface LayoutSnap {
  sheetId: string;
  sheetName: string;
  /** [zero-based index, size in points] */
  columns: Array<[number, number]>;
  rows: Array<[number, number]>;
}

/** An Excel table on a sheet the tool touched: writing next to one can grow it, and undo must shrink it back. */
export interface TableSnap {
  sheetId: string;
  id: string;
  name: string;
  row: number;
  col: number;
  rows: number;
  cols: number;
}

export interface Photo {
  blocks: Block[];
  layouts: LayoutSnap[];
  tables: TableSnap[];
}

/** One 32-bit hash per cell (formula, number format and style), plus keys for merges and sizes. */
export interface Fingerprint {
  cells: number[][];
  merged: string[];
  layouts: string[];
  /** One key per photographed table, in the same order. */
  tables: string[];
}

interface Flags {
  style: boolean;
  merged: boolean;
}

/* ----- hashing and styles (pure) ----- */

function fnv(text: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return h >>> 0;
}

const stripOdata = (value: unknown): unknown =>
  Array.isArray(value)
    ? value.map(stripOdata)
    : value && typeof value === "object"
      ? Object.fromEntries(Object.entries(value).filter(([key]) => key !== "@odata.type").map(([key, v]) => [key, stripOdata(v)]))
      : value;

const BORDER_SIDES = ["top", "bottom", "left", "right"] as const;

/** One cell's look as stable JSON: only what getCellProperties/setCellProperties both understand, no empty borders. */
export function canonicalStyle(props: unknown): string {
  const format = ((stripOdata(props) as { format?: Record<string, any> } | undefined)?.format ?? {}) as Record<string, any>;
  const borders: Record<string, unknown> = {};
  for (const side of BORDER_SIDES) {
    const border = format.borders?.[side];
    if (border?.style) borders[side] = { style: border.style, color: border.color, weight: border.weight };
  }
  return JSON.stringify({
    font: format.font,
    fill: format.fill?.color ?? "",
    horizontalAlignment: format.horizontalAlignment,
    verticalAlignment: format.verticalAlignment,
    wrapText: format.wrapText,
    indentLevel: format.indentLevel,
    borders,
  });
}

function styleGrid(props: unknown[][]): { palette: string[]; grid: number[][] } {
  const palette: string[] = [];
  const index = new Map<string, number>();
  const grid = props.map((row) =>
    row.map((cell) => {
      const key = canonicalStyle(cell);
      let at = index.get(key);
      if (at === undefined) {
        at = palette.length;
        palette.push(key);
        index.set(key, at);
      }
      return at;
    }),
  );
  return { palette, grid };
}

export function cellHash(formula: unknown, numberFormat: string, style: string | undefined): number {
  return fnv(`${typeof formula}\u0001${String(formula)}\u0001${numberFormat}\u0001${style ?? ""}`);
}

function blockHashes(block: Block): number[] {
  const out: number[] = [];
  for (let r = 0; r < block.rows; r++) {
    for (let c = 0; c < block.cols; c++) {
      const style = block.styles ? block.styles.palette[block.styles.grid[r]![c]!] : undefined;
      out.push(cellHash(block.formulas[r]![c], block.numberFormat[r]![c]!, style));
    }
  }
  return out;
}

const mergedKey = (block: Block) => (block.merged ?? []).map((m) => m.join(",")).sort().join(";");
const layoutKey = (layout: LayoutSnap) => `${layout.sheetId}|${layout.columns.map((c) => c.join(":")).join(",")}|${layout.rows.map((r) => r.join(":")).join(",")}`;

const tableKey = (table: TableSnap | undefined, id: string) => (table ? `${id}|${table.row},${table.col},${table.rows},${table.cols}` : `${id}|gone`);

export function fingerprintOf(photo: Photo): Fingerprint {
  return {
    cells: photo.blocks.map(blockHashes),
    merged: photo.blocks.map(mergedKey),
    layouts: photo.layouts.map(layoutKey),
    tables: photo.tables.map((table) => tableKey(table, table.id)),
  };
}

export const sameFingerprint = (a: Fingerprint, b: Fingerprint) =>
  JSON.stringify(a.cells) === JSON.stringify(b.cells) &&
  a.merged.join("|") === b.merged.join("|") &&
  a.layouts.join("|") === b.layouts.join("|") &&
  a.tables.join("|") === b.tables.join("|");

export function photoCells(photo: Photo): number {
  return photo.blocks.reduce((sum, block) => sum + block.rows * block.cols, 0);
}

/** Text that Excel would turn into a number, date, boolean or formula if it were typed into a General cell. */
export function needsTextPrefix(text: string): boolean {
  if (text === "") return false;
  return /^[=+\-@.$€£¥#('0-9]/.test(text) || /^(true|false)$/i.test(text) || /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s\-/]*\d/i.test(text);
}

/* ----- reading Excel ----- */

interface Geometry {
  sheetId: string;
  sheetName: string;
  row: number;
  col: number;
  rows: number;
  cols: number;
}

const CELL_QUERY = (() => {
  const side = { style: true, color: true, weight: true };
  return {
    format: {
      fill: { color: true },
      font: { bold: true, italic: true, color: true, name: true, size: true, underline: true, strikethrough: true },
      horizontalAlignment: true,
      verticalAlignment: true,
      wrapText: true,
      indentLevel: true,
      borders: { top: side, bottom: side, left: side, right: side },
    },
  } as unknown as Excel.CellPropertiesLoadOptions;
})();

const HUGE = { rows: 100_000, cols: 2_000 };

async function resolveGeometries(context: Excel.RequestContext, targets: UndoTarget[], limitCells: number, valuesOnly: boolean): Promise<Geometry[]> {
  type Pending = { sheet: Excel.Worksheet; range: Excel.Range | null; used: boolean };
  const pending: Pending[] = [];
  const sheets = context.workbook.worksheets;
  const wantsAll = targets.some((target) => "allSheets" in target);
  if (wantsAll) {
    sheets.load("items/id,items/name,items/visibility");
    await context.sync();
    for (const sheet of sheets.items) if (sheet.visibility === "Visible") pending.push({ sheet, range: null, used: true });
  }
  for (const target of targets) {
    if ("allSheets" in target) continue;
    const sheet = resolveSheet(context, target.sheet);
    sheet.load("id,name");
    let range = sheet.getRange(target.address);
    if (target.copySource) {
      const source = sheet.getRange(target.copySource);
      source.load("rowCount,columnCount");
      range.load("rowCount,columnCount");
      await context.sync();
      range = range.getCell(0, 0).getResizedRange(Math.max(range.rowCount, source.rowCount) - 1, Math.max(range.columnCount, source.columnCount) - 1);
    }
    if (target.grow) range = range.getCell(0, 0).getResizedRange(target.grow.rows - 1, target.grow.columns - 1);
    range.load("rowIndex,columnIndex,rowCount,columnCount");
    pending.push({ sheet, range, used: false });
  }
  await context.sync();

  // Whole rows/columns are narrowed to what the sheet actually uses.
  const narrowed = pending.map((item) => {
    const needsUsed = item.used || (item.range && (item.range.rowCount > HUGE.rows || item.range.columnCount > HUGE.cols));
    if (!needsUsed) return { item, range: item.range!, narrowed: null as Excel.Range | null };
    const used = item.sheet.getUsedRangeOrNullObject(valuesOnly);
    const inside = item.range ? item.range.getIntersectionOrNullObject(used) : used;
    inside.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    return { item, range: item.range, narrowed: inside };
  });
  await context.sync();

  const out: Geometry[] = [];
  let total = 0;
  for (const { item, range, narrowed: inside } of narrowed) {
    if (inside?.isNullObject) continue;
    const r = (inside ?? range)!;
    out.push({ sheetId: item.sheet.id, sheetName: item.sheet.name, row: r.rowIndex, col: r.columnIndex, rows: r.rowCount, cols: r.columnCount });
    total += r.rowCount * r.columnCount;
    if (total > limitCells) throw new TooLarge(`this change touches more than ${limitCells.toLocaleString()} cells`);
  }
  return out;
}

async function readBlocks(context: Excel.RequestContext, geometries: Geometry[], flags: Flags): Promise<Block[]> {
  const queued = geometries.map((g) => {
    const sheet = context.workbook.worksheets.getItem(g.sheetId);
    const range = sheet.getRangeByIndexes(g.row, g.col, g.rows, g.cols);
    range.load("formulas,valueTypes,numberFormat");
    const props = flags.style ? range.getCellProperties(CELL_QUERY) : null;
    let merged: Excel.RangeAreas | null = null;
    if (flags.merged) {
      merged = range.getMergedAreasOrNullObject();
      merged.load("isNullObject");
      merged.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
    }
    return { g, range, props, merged };
  });
  await context.sync();
  return queued.map(({ g, range, props, merged }) => ({
    ...g,
    formulas: range.formulas,
    valueTypes: range.valueTypes as Grid<string>,
    numberFormat: range.numberFormat as Grid<string>,
    ...(props ? { styles: styleGrid(props.value as unknown[][]) } : {}),
    ...(merged
      ? { merged: merged.isNullObject ? [] : merged.areas.items.map((a): [number, number, number, number] => [a.rowIndex, a.columnIndex, a.rowCount, a.columnCount]) }
      : {}),
  }));
}

type LayoutWanted = { sheetId: string; sheetName: string; columns: number[]; rows: number[] };

/** Which columns/rows a resize may change. A range with no address means the sheet's used range. */
async function resolveLayouts(context: Excel.RequestContext, spec: LayoutSpec): Promise<LayoutWanted[]> {
  const sheet = resolveSheet(context, spec.sheet);
  sheet.load("id,name");
  let columns: [number, number] | undefined = spec.columns;
  let rows: [number, number] | undefined = spec.rows;
  if (!columns && !rows) {
    const range = spec.address ? sheet.getRange(spec.address) : sheet.getUsedRangeOrNullObject(false);
    range.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();
    if (range.isNullObject) return [];
    // A whole-column reference spans every row; only the dimension that was asked for counts.
    if (spec.dimensions.includes("width")) columns = [range.columnIndex + 1, range.columnIndex + range.columnCount];
    if (spec.dimensions.includes("height")) rows = [range.rowIndex + 1, range.rowIndex + range.rowCount];
  } else {
    await context.sync();
  }
  const columnList = spec.dimensions.includes("width") && columns ? span(columns) : [];
  const rowList = spec.dimensions.includes("height") && rows ? span(rows) : [];
  if (columnList.length > PHOTO_LIMITS.columns || rowList.length > PHOTO_LIMITS.rows) {
    throw new TooLarge("this change resizes too many columns or rows");
  }
  return [{ sheetId: sheet.id, sheetName: sheet.name, columns: columnList, rows: rowList }];
}

const span = ([from, to]: [number, number]) => Array.from({ length: Math.max(0, Math.min(to - from + 1, 1_000_000)) }, (_, i) => from - 1 + i);

async function readLayouts(context: Excel.RequestContext, wanted: LayoutWanted[]): Promise<LayoutSnap[]> {
  const queued = wanted.map((w) => {
    const sheet = context.workbook.worksheets.getItem(w.sheetId);
    const columns = w.columns.map((index) => {
      const format = sheet.getRangeByIndexes(0, index, 1, 1).getEntireColumn().format;
      format.load("columnWidth");
      return { index, format };
    });
    const rows = w.rows.map((index) => {
      const format = sheet.getRangeByIndexes(index, 0, 1, 1).getEntireRow().format;
      format.load("rowHeight");
      return { index, format };
    });
    return { w, columns, rows };
  });
  await context.sync();
  return queued.map(({ w, columns, rows }) => ({
    sheetId: w.sheetId,
    sheetName: w.sheetName,
    columns: columns.map((c): [number, number] => [c.index, c.format.columnWidth]),
    rows: rows.map((r): [number, number] => [r.index, r.format.rowHeight]),
  }));
}

/** Every table on the given sheets, with where it sits now. */
async function readTables(context: Excel.RequestContext, sheetIds: string[]): Promise<TableSnap[]> {
  const sheets = sheetIds.map((id) => {
    const sheet = context.workbook.worksheets.getItem(id);
    sheet.tables.load("items/id,items/name");
    return { id, sheet };
  });
  await context.sync();
  const found = sheets.flatMap(({ id, sheet }) =>
    sheet.tables.items.map((table) => {
      const range = table.getRange();
      range.load("rowIndex,columnIndex,rowCount,columnCount");
      return { id, table, range };
    }),
  );
  await context.sync();
  return found.map(({ id, table, range }) => ({ sheetId: id, id: table.id, name: table.name, row: range.rowIndex, col: range.columnIndex, rows: range.rowCount, cols: range.columnCount }));
}

const sheetsOf = (blocks: Block[], layouts: LayoutSnap[]) => [...new Set([...blocks.map((b) => b.sheetId), ...layouts.map((l) => l.sheetId)])];

type CellsPlan = Extract<UndoPlan, { kind: "cells" }>;

/** Photograph everything `plan` may change. Throws TooLarge when it is too big to keep. */
export async function photograph(plan: CellsPlan): Promise<Photo> {
  return Excel.run(async (context) => {
    const geometries = await resolveGeometries(context, plan.targets, plan.style ? PHOTO_LIMITS.styledCells : PHOTO_LIMITS.cells, !plan.style);
    const blocks = await readBlocks(context, geometries, { style: plan.style, merged: !!plan.merged });
    const wanted = plan.layout ? await resolveLayouts(context, plan.layout) : [];
    const layouts = await readLayouts(context, wanted);
    return { blocks, layouts, tables: await readTables(context, sheetsOf(blocks, layouts)) };
  });
}

/** What the photographed cells look like now (same cells, same flags as the photo). */
export async function fingerprintNow(photo: Photo): Promise<Fingerprint> {
  return Excel.run(async (context) => {
    const geometries: Geometry[] = photo.blocks.map(({ sheetId, sheetName, row, col, rows, cols }) => ({ sheetId, sheetName, row, col, rows, cols }));
    const flags = { style: photo.blocks.some((b) => b.styles), merged: photo.blocks.some((b) => b.merged) };
    const blocks = await readBlocks(context, geometries, flags);
    const layouts = await readLayouts(
      context,
      photo.layouts.map((l) => ({ sheetId: l.sheetId, sheetName: l.sheetName, columns: l.columns.map(([i]) => i), rows: l.rows.map(([i]) => i) })),
    );
    const now = await readTables(context, sheetsOf(photo.blocks, photo.layouts));
    const tables = photo.tables.map((before) => now.find((t) => t.id === before.id && t.sheetId === before.sheetId));
    return { ...fingerprintOf({ blocks, layouts, tables: [] }), tables: photo.tables.map((before, i) => tableKey(tables[i], before.id)) };
  });
}

/* ----- putting cells back ----- */

export interface RestoreReport {
  restoredCells: number;
  /** Cells someone changed after the tool ran; left as they are. */
  conflicts: Array<{ sheet: string; address: string }>;
  conflictCount: number;
  /** Cells Excel stored differently from the photo even after a retry. */
  unrestored: Array<{ sheet: string; address: string }>;
  unrestoredCount: number;
  /** Cells on a sheet that has since been deleted. */
  missingCells: number;
  notes: string[];
}

const ADDRESS_PREVIEW = 20;

function a1(row: number, col: number): string {
  let letters = "";
  for (let v = col + 1; v > 0; v = Math.floor((v - 1) / 26)) letters = String.fromCharCode(65 + ((v - 1) % 26)) + letters;
  return `${letters}${row + 1}`;
}

const same = (a: unknown, b: unknown) => a === b;

function textRestorable(before: Block, r: number, c: number, force: boolean): unknown {
  const formula = before.formulas[r]![c];
  if (typeof formula !== "string" || before.valueTypes[r]![c] !== "String" || formula.startsWith("=")) return formula;
  if (before.numberFormat[r]![c] === "@") return formula;
  return force || needsTextPrefix(formula) ? `'${formula}` : formula;
}

/**
 * Put one block back. Cells whose fingerprint differs from `after` were changed by someone else since the tool ran:
 * they keep their current content. The write is checked by reading it back; strings Excel coerced (00123 -> 123) are
 * retried with a text prefix.
 */
async function restoreBlock(context: Excel.RequestContext, before: Block, after: number[], mergedAfter: string, report: RestoreReport, done: Set<string>): Promise<void> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(before.sheetId);
  await context.sync();
  if (sheet.isNullObject) {
    report.missingCells += before.rows * before.cols;
    return;
  }
  const [current] = await readBlocks(context, [before], { style: !!before.styles, merged: !!before.merged });
  const currentHashes = blockHashes(current!);

  const wanted: Block = { ...before, formulas: before.formulas.map((row) => [...row]), numberFormat: before.numberFormat.map((row) => [...row]) };
  const keep: boolean[][] = before.formulas.map((row) => row.map(() => false));
  const styleOf = (block: Block, r: number, c: number) => block.styles!.palette[block.styles!.grid[r]![c]!]!;
  const palette = before.styles ? [...before.styles.palette] : [];
  const grid = before.styles ? before.styles.grid.map((row) => [...row]) : [];
  let conflicts = 0;
  let skipped = 0;
  for (let r = 0; r < before.rows; r++) {
    for (let c = 0; c < before.cols; c++) {
      const key = `${before.sheetId}|${before.row + r}|${before.col + c}`;
      const mine = currentHashes[r * before.cols + c] === after[r * before.cols + c];
      // Two photos of one change can overlap; the first to be restored (or flagged) settles the cell.
      const settled = done.has(key);
      done.add(key);
      if (mine && !settled) continue;
      keep[r]![c] = true;
      if (settled) skipped++;
      else conflicts++;
      wanted.formulas[r]![c] = current!.formulas[r]![c];
      wanted.valueTypes[r]![c] = current!.valueTypes[r]![c]!;
      wanted.numberFormat[r]![c] = current!.numberFormat[r]![c]!;
      if (before.styles) {
        const style = styleOf(current!, r, c);
        let at = palette.indexOf(style);
        if (at < 0) at = palette.push(style) - 1;
        grid[r]![c] = at;
      }
      if (!settled && report.conflicts.length < ADDRESS_PREVIEW) report.conflicts.push({ sheet: before.sheetName, address: a1(before.row + r, before.col + c) });
    }
  }
  report.conflictCount += conflicts;
  if (conflicts + skipped === before.rows * before.cols) return;

  const range = sheet.getRangeByIndexes(before.row, before.col, before.rows, before.cols);
  const touchMerges = !!before.merged && mergedKeyOf(current!) === mergedAfter;
  if (before.merged && !touchMerges) report.notes.push("merged cells were changed after the tool ran and were left alone");
  if (touchMerges) range.unmerge();
  if (before.styles) {
    range.format.fill.clear();
    for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const) {
      range.format.borders.getItem(edge).style = "None";
    }
  }
  const write = async (force: Set<string> | "none") => {
    range.numberFormat = wanted.numberFormat;
    range.formulas = wanted.formulas.map((row, r) =>
      row.map((_, c) => (keep[r]![c] ? wanted.formulas[r]![c] : textRestorable(before, r, c, force !== "none" && force.has(`${r},${c}`)))),
    ) as any;
    await context.sync();
  };
  await write("none");

  if (before.styles) {
    const props = grid.map((row) =>
      row.map((index) => {
        const style = JSON.parse(palette[index]!) as Record<string, any>;
        const format: Record<string, unknown> = { ...style };
        if (!style.fill) delete format.fill; else format.fill = { color: style.fill };
        if (!Object.keys(style.borders ?? {}).length) delete format.borders;
        for (const key of Object.keys(format)) if (format[key] === undefined) delete format[key];
        return { format };
      }),
    );
    range.setCellProperties(props as any);
    await context.sync();
  }
  if (touchMerges) {
    for (const [row, col, rows, cols] of before.merged!) sheet.getRangeByIndexes(row, col, rows, cols).merge(false);
    await context.sync();
  }

  // Read back: anything Excel stored differently gets one retry with a text prefix.
  const verify = async () => {
    const [now] = await readBlocks(context, [before], { style: false, merged: false });
    const bad: Array<[number, number]> = [];
    for (let r = 0; r < before.rows; r++) {
      for (let c = 0; c < before.cols; c++) {
        if (keep[r]![c]) continue;
        const touchedByMerge = touchMerges && before.merged!.some(([mr, mc, mrows, mcols]) => before.row + r >= mr && before.row + r < mr + mrows && before.col + c >= mc && before.col + c < mc + mcols && !(before.row + r === mr && before.col + c === mc));
        if (touchedByMerge) continue;
        if (!same(now!.formulas[r]![c], before.formulas[r]![c]) || now!.valueTypes[r]![c] !== before.valueTypes[r]![c]) bad.push([r, c]);
      }
    }
    return bad;
  };
  let bad = await verify();
  if (bad.length > 0 && !touchMerges) {
    await write(new Set(bad.map(([r, c]) => `${r},${c}`)));
    bad = await verify();
  }
  for (const [r, c] of bad) {
    report.unrestoredCount++;
    if (report.unrestored.length < ADDRESS_PREVIEW) report.unrestored.push({ sheet: before.sheetName, address: a1(before.row + r, before.col + c) });
  }
  report.restoredCells += before.rows * before.cols - conflicts - skipped - bad.length;
}

function mergedKeyOf(block: Block): string {
  return mergedKey(block);
}

/** Put every block and size back. `after` is the fingerprint taken right after the tool ran. */
export async function restorePhoto(photo: Photo, after: Fingerprint): Promise<RestoreReport> {
  const report: RestoreReport = { restoredCells: 0, conflicts: [], conflictCount: 0, unrestored: [], unrestoredCount: 0, missingCells: 0, notes: [] };
  await Excel.run(async (context) => {
    const done = new Set<string>();
    for (let i = 0; i < photo.blocks.length; i++) {
      await restoreBlock(context, photo.blocks[i]!, after.cells[i]!, after.merged[i] ?? "", report, done);
    }
    for (let i = 0; i < photo.layouts.length; i++) {
      const layout = photo.layouts[i]!;
      const now = (await readLayouts(context, [{ sheetId: layout.sheetId, sheetName: layout.sheetName, columns: layout.columns.map(([n]) => n), rows: layout.rows.map(([n]) => n) }]))[0]!;
      if (layoutKey(now) !== after.layouts[i]) {
        report.notes.push(`column widths or row heights on "${layout.sheetName}" were changed after the tool ran and were left alone`);
        continue;
      }
      const sheet = context.workbook.worksheets.getItemOrNullObject(layout.sheetId);
      await context.sync();
      if (sheet.isNullObject) continue;
      for (const [index, width] of layout.columns) sheet.getRangeByIndexes(0, index, 1, 1).getEntireColumn().format.columnWidth = width;
      for (const [index, height] of layout.rows) sheet.getRangeByIndexes(index, 0, 1, 1).getEntireRow().format.rowHeight = height;
      await context.sync();
    }
    // Writing next to a table can make Excel grow it; give it its old size back, unless someone resized it since.
    for (let i = 0; i < photo.tables.length; i++) {
      const before = photo.tables[i]!;
      const sheet = context.workbook.worksheets.getItemOrNullObject(before.sheetId);
      await context.sync();
      if (sheet.isNullObject) continue;
      const table = sheet.tables.getItemOrNullObject(before.id);
      table.load("isNullObject");
      await context.sync();
      if (table.isNullObject) continue;
      const range = table.getRange();
      range.load("rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      const now: TableSnap = { ...before, row: range.rowIndex, col: range.columnIndex, rows: range.rowCount, cols: range.columnCount };
      if (tableKey(now, before.id) === tableKey(before, before.id)) continue;
      if (tableKey(now, before.id) !== after.tables[i]) {
        report.notes.push(`the table "${before.name}" was resized after the tool ran and was left alone`);
        continue;
      }
      table.resize(sheet.getRangeByIndexes(before.row, before.col, before.rows, before.cols));
      await context.sync();
    }
  });
  return report;
}

/* ----- inserted rows ----- */

export interface InsertedRows {
  sheetId: string;
  sheetName: string;
  /** 1-based first row and count. */
  row: number;
  count: number;
}

export async function locateInsert(sheetRef: Parameters<typeof resolveSheet>[1], row: number, count: number): Promise<InsertedRows> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, sheetRef);
    sheet.load("id,name");
    await context.sync();
    return { sheetId: sheet.id, sheetName: sheet.name, row, count };
  });
}

/** Delete the inserted rows again, but only while they are still empty. Returns a refusal message, or null when done. */
export async function removeInsertedRows(inserted: InsertedRows): Promise<string | null> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(inserted.sheetId);
    await context.sync();
    if (sheet.isNullObject) return `the sheet "${inserted.sheetName}" no longer exists`;
    const rows = sheet.getRange(`${inserted.row}:${inserted.row + inserted.count - 1}`);
    const used = rows.getUsedRangeOrNullObject(true);
    used.load("isNullObject");
    await context.sync();
    if (!used.isNullObject) return `the rows inserted on "${inserted.sheetName}" now hold data`;
    rows.delete(Excel.DeleteShiftDirection.up);
    await context.sync();
    return null;
  });
}
