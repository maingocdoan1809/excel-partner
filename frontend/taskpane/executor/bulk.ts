import { isFormulaString } from "./common.js";

/**
 * Bulk writing that keeps Excel responsive. Excel does a write on its main thread, so the only lever is the size of
 * each round trip: writing formulas costs about 0.15 ms per cell, copying them with `copyFrom` about 0.01 ms, values
 * in between. Batches are therefore sized by measured time (`BULK.targetMs`), not by a fixed number of cells.
 */

/** Tunable from tests. */
export const BULK = {
  /** Time one round trip should take; keeps a pause the person can still feel to a fraction of a second. */
  targetMs: 400,
  /** After the person last moved in the workbook, wait this long (at most `maxWaitMs`) before the next batch. */
  idleMs: 300,
  maxWaitMs: 2_000,
  now: (): number => (typeof performance !== "undefined" ? performance.now() : Date.now()),
  /** Lets Excel and the task pane handle their own events between two batches. */
  pause: (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms)),
};

let lastUserActivity = Number.NEGATIVE_INFINITY;

/** Call when the person acts in the workbook (for example moves the selection): the next batch then waits a moment. */
export function noteUserActivity(): void {
  lastUserActivity = BULK.now();
}

/** Pause between two batches; longer while the person is working in the workbook. */
async function breathe(tookMs: number): Promise<void> {
  await BULK.pause(Math.min(120, Math.max(15, Math.round(tookMs * 0.3))));
  for (let waited = 0; waited < BULK.maxWaitMs && BULK.now() - lastUserActivity < BULK.idleMs; waited += 100) await BULK.pause(100);
}

export type BulkKind = "values" | "formulas" | "fill";

/** First batch, smallest and largest batch, in cells. */
const LIMITS: Record<BulkKind, { start: number; min: number; max: number }> = {
  values: { start: 5_000, min: 500, max: 50_000 },
  formulas: { start: 1_000, min: 200, max: 10_000 },
  fill: { start: 5_000, min: 500, max: 50_000 },
};

export interface BulkStats {
  batches: number;
  totalMs: number;
  slowestMs: number;
}

export const emptyStats = (): BulkStats => ({ batches: 0, totalMs: 0, slowestMs: 0 });

export function mergeStats(...all: BulkStats[]): BulkStats {
  return all.reduce((sum, one) => ({ batches: sum.batches + one.batches, totalMs: sum.totalMs + one.totalMs, slowestMs: Math.max(sum.slowestMs, one.slowestMs) }), emptyStats());
}

/** Whether a block of this many cells needs more than the first batch. */
export const needsBatches = (kind: BulkKind, cells: number): boolean => cells > LIMITS[kind].start;

/**
 * Run `step(firstRow, rowCount)` over `rows` rows, one `sync` per batch. The next batch grows or shrinks so that it
 * takes about `BULK.targetMs`, and there is a short pause between batches.
 */
export async function inBatches(
  context: { sync(): Promise<unknown>; application?: { suspendScreenUpdatingUntilNextSync?(): void } },
  kind: BulkKind,
  rows: number,
  columns: number,
  step: (firstRow: number, rowCount: number) => void,
): Promise<BulkStats> {
  const limit = LIMITS[kind];
  const stats = emptyStats();
  let cells = limit.start;
  for (let start = 0; start < rows; ) {
    const count = Math.min(rows - start, Math.max(1, Math.floor(cells / Math.max(1, columns))));
    const began = BULK.now();
    // Ends by itself at the next sync, so the screen is only frozen for the length of one batch.
    context.application?.suspendScreenUpdatingUntilNextSync?.();
    step(start, count);
    await context.sync();
    const took = Math.max(0, BULK.now() - began);
    if (pausedFrom !== null) remember(pausedFrom);
    stats.batches++;
    stats.totalMs += took;
    stats.slowestMs = Math.max(stats.slowestMs, took);
    start += count;
    const factor = Math.min(2, Math.max(0.25, BULK.targetMs / Math.max(took, 1)));
    cells = Math.round(Math.min(limit.max, Math.max(limit.min, count * Math.max(1, columns) * factor)));
    if (start < rows) await breathe(took);
  }
  return stats;
}

/**
 * Survives a crash of the task pane, so a later one can put Excel's calculation mode back. The pane that is writing
 * refreshes the time on every batch, so another pane only takes over a mode nobody has touched for a while.
 */
const MODE_KEY = "office-partner.calculation-mode-to-restore";
const STALE_AFTER_MS = 60_000;
let pausedFrom: string | null = null;

const remember = (mode: string | null): void => {
  try {
    if (mode === null) localStorage.removeItem(MODE_KEY);
    else localStorage.setItem(MODE_KEY, JSON.stringify({ mode, at: Date.now() }));
  } catch {
    /* storage can be blocked; the finally block still restores the mode in the normal case */
  }
};

/**
 * If an earlier task pane stopped while calculation was paused, put the mode back. Returns the mode it restored.
 * Only acts when the record is old and Excel is still on manual, so a write running in another pane, or a mode the
 * person chose themselves, is left alone.
 */
export async function recoverCalculationMode(): Promise<string | null> {
  let saved: { mode: string; at: number } | null = null;
  try {
    const raw = localStorage.getItem(MODE_KEY);
    saved = raw ? (JSON.parse(raw) as { mode: string; at: number }) : null;
  } catch {
    return null;
  }
  if (!saved || typeof saved.mode !== "string") return null;
  if (Date.now() - Number(saved.at) < STALE_AFTER_MS) return null;
  const { mode } = saved;
  const restored = await Excel.run(async (context) => {
    context.application.load("calculationMode");
    await context.sync();
    if (context.application.calculationMode !== Excel.CalculationMode.manual || mode === Excel.CalculationMode.manual) return null;
    context.application.calculationMode = mode as Excel.CalculationMode;
    await context.sync();
    return mode;
  });
  remember(null);
  return restored;
}

/** Whether to pause recalculation is the caller's call; the pause itself costs a few round trips, so small writes skip it. */
export function manualCalculationIf<T>(context: Excel.RequestContext, needed: boolean, work: () => Promise<T>): Promise<T> {
  return needed ? withManualCalculation(context, work) : work();
}

/** The `timing` block tool results carry, so a slow run can be traced to its slowest round trip. */
export const timingOf = (stats: BulkStats, extra: Record<string, number> = {}) => ({ batches: stats.batches, totalMs: Math.round(stats.totalMs), slowestBatchMs: Math.round(stats.slowestMs), ...extra });

/**
 * Run `work` with recalculation off, so a write made of several batches does not recalculate the workbook after each
 * one. Excel's previous mode is put back in `finally` (which recalculates once), even when the work fails.
 */
export async function withManualCalculation<T>(context: Excel.RequestContext, work: () => Promise<T>): Promise<T> {
  const app = context.application;
  app.load("calculationMode");
  await context.sync();
  const saved = app.calculationMode;
  if (saved === Excel.CalculationMode.manual) return work();
  pausedFrom = saved;
  remember(saved);
  app.calculationMode = Excel.CalculationMode.manual;
  await context.sync();
  let failure: unknown;
  try {
    return await work();
  } catch (error) {
    failure = error;
    throw error;
  } finally {
    try {
      app.calculationMode = saved;
      await context.sync();
      pausedFrom = null;
      remember(null);
    } catch (restoreError) {
      // The record stays, so a later pane puts the mode back; a failure of the work itself is the one to report.
      if (failure === undefined) throw restoreError;
    }
  }
}

/** A range with more cells than this is worked on in row batches instead of one operation. */
export const LARGE_RANGE_CELLS = 20_000;

/** Apply `operation` to consecutive row slices of `range` (same columns), in time-adaptive batches. */
export function overRows(
  context: Excel.RequestContext,
  range: Excel.Range,
  rows: number,
  columns: number,
  kind: BulkKind,
  operation: (slice: Excel.Range, rowCount: number) => void,
): Promise<BulkStats> {
  return inBatches(context, kind, rows, columns, (start, count) => operation(range.getCell(start, 0).getResizedRange(count - 1, columns - 1), count));
}

/**
 * `destination.copyFrom(source)` in row batches. A source with several rows is repeated down the destination, so each
 * batch covers whole repeats: the relative references of every copy shift exactly as in one big copy.
 */
export function copyInBatches(
  context: Excel.RequestContext,
  source: Excel.Range,
  sourceRows: number,
  destination: Excel.Range,
  destinationRows: number,
  destinationColumns: number,
  type: Excel.RangeCopyType,
): Promise<BulkStats> {
  const repeats = Math.ceil(destinationRows / Math.max(1, sourceRows));
  return inBatches(context, "fill", repeats, Math.max(1, sourceRows) * destinationColumns, (firstRepeat, repeatCount) => {
    const firstRow = firstRepeat * sourceRows;
    const rowCount = Math.min(destinationRows - firstRow, repeatCount * sourceRows);
    // A last, shorter batch copies only as many source rows as it has room for; a destination smaller than the
    // source would otherwise be replaced by the whole source and spill past the end.
    const part = rowCount < sourceRows ? source.getResizedRange(rowCount - sourceRows, 0) : source;
    destination.getCell(firstRow, 0).getResizedRange(rowCount - 1, destinationColumns - 1).copyFrom(part, type);
  });
}

/** Excel can rewrite a formula's spelling (case, spaces); compare without those. */
const sameFormula = (a: unknown, b: unknown): boolean => String(a).replace(/\s+/g, "").toUpperCase() === String(b).replace(/\s+/g, "").toUpperCase();
/** Below this many rows, writing every formula is not worth a fill-and-check. */
const FILL_MIN_ROWS = 20;
/** Cells re-read per round trip when checking a filled column. */
const CHECK_CELLS = 50_000;

export interface GridWrite extends BulkStats {
  /** Columns whose formulas were written once and filled down instead of written row by row. */
  filledColumns: number;
}

/**
 * Write a 2D array whose top-left cell is `topLeft`, in time-adaptive batches with calculation paused when the write
 * is big. A column that holds a formula in every row is written in its first row and copied down: Excel copies
 * formulas about 15 times faster than it accepts them one by one. Each filled column is then read back, and any row
 * whose formula is not what was asked for is written explicitly, so the result is the same either way.
 * `null` means "leave the cell alone". Offered to `execute_office_js` as `partner.writeGrid`.
 */
export async function writeGrid(context: Excel.RequestContext, topLeft: Excel.Range, grid: Array<Array<unknown>>): Promise<GridWrite> {
  const anchor = topLeft.getCell(0, 0);
  const rows = grid.length;
  const columns = grid.reduce((max, row) => Math.max(max, row.length), 0);
  if (rows === 0 || columns === 0) return { ...emptyStats(), filledColumns: 0 };

  const fillable = new Set<number>();
  if (rows >= FILL_MIN_ROWS) {
    for (let c = 0; c < columns; c++) if (grid.every((row) => isFormulaString(row[c]))) fillable.add(c);
  }
  // Rows sent in the batches: a filled column only sends its first row.
  const rowToSend = (r: number): unknown[] => {
    const row = grid[r]!;
    if (row.length === columns && (fillable.size === 0 || r === 0)) return row;
    return Array.from({ length: columns }, (_, c) => (c >= row.length || (r > 0 && fillable.has(c)) ? null : row[c]));
  };
  const hasFormulas = grid.some((row, r) => rowToSend(r).some(isFormulaString));
  const kind: BulkKind = hasFormulas ? "formulas" : "values";

  return manualCalculationIf(context, needsBatches(kind, rows * columns), async () => {
    let stats = await inBatches(context, kind, rows, columns, (start, count) => {
      const batch = [];
      for (let r = start; r < start + count; r++) batch.push(rowToSend(r));
      anchor.getOffsetRange(start, 0).getResizedRange(count - 1, columns - 1).values = batch;
    });

    // Runs of neighbouring filled columns are handled together.
    const runs: Array<[number, number]> = [];
    for (const c of [...fillable].sort((a, b) => a - b)) {
      const last = runs[runs.length - 1];
      if (last && last[1] === c - 1) last[1] = c;
      else runs.push([c, c]);
    }
    for (const [first, last] of runs) {
      const width = last - first + 1;
      const seed = anchor.getOffsetRange(0, first).getResizedRange(0, width - 1);
      stats = mergeStats(
        stats,
        await inBatches(context, "fill", rows - 1, width, (start, count) => {
          anchor.getOffsetRange(1 + start, first).getResizedRange(count - 1, width - 1).copyFrom(seed, Excel.RangeCopyType.formulas);
        }),
      );

      // Read back, in chunks, and rewrite only the rows that differ.
      const wrong: number[] = [];
      const chunkRows = Math.max(1, Math.floor(CHECK_CELLS / width));
      for (let start = 1; start < rows; start += chunkRows) {
        const count = Math.min(chunkRows, rows - start);
        const block = anchor.getOffsetRange(start, first).getResizedRange(count - 1, width - 1);
        block.load("formulas");
        await context.sync();
        block.formulas.forEach((formulas, i) => {
          if (formulas.some((formula, c) => !sameFormula(formula, grid[start + i]![first + c]))) wrong.push(start + i);
        });
      }
      if (wrong.length > 0) {
        stats = mergeStats(
          stats,
          await inBatches(context, "formulas", wrong.length, width, (from, count) => {
            for (const row of wrong.slice(from, from + count)) {
              anchor.getOffsetRange(row, first).getResizedRange(0, width - 1).values = [grid[row]!.slice(first, last + 1)];
            }
          }),
        );
      }
    }
    return { ...stats, filledColumns: fillable.size };
  });
}

/**
 * Copy `source` over `destination` (repeating it down) in time-adaptive batches, with calculation paused when it is
 * big. For `execute_office_js` code that fills a block from a sample row. `type` is formulas, all, values or formats.
 */
export async function fillDown(context: Excel.RequestContext, source: Excel.Range, destination: Excel.Range, type: "formulas" | "all" | "values" | "formats" = "formulas"): Promise<BulkStats> {
  source.load("rowCount");
  destination.load("rowCount,columnCount");
  await context.sync();
  const cells = destination.rowCount * destination.columnCount;
  const copyType = Excel.RangeCopyType[type];
  return manualCalculationIf(context, needsBatches("fill", cells), () =>
    copyInBatches(context, source, source.rowCount, destination, destination.rowCount, destination.columnCount, copyType),
  );
}

/** The helpers `execute_office_js` code receives as `partner`. */
export const bulkHelpers = { manualCalculation: withManualCalculation, writeGrid, fillDown, inBatches };
