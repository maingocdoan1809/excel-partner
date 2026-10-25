import {
  EXECUTOR_CAPS,
  numberToColumn,
  type CheckFormulaConsistencyInput,
  type FindErrorsInput,
  type SheetRef,
} from "../../../protocol/src/index.js";
import { isErrorValue, resolveSheet } from "./common.js";

/** One block of cells handed to a scan visitor. Indexes are zero-based and absolute on the sheet. */
export interface ScanChunk {
  sheet: string;
  rowIndex: number;
  columnIndex: number;
  rows: number;
  columns: number;
  values: unknown[][];
  formulas: unknown[][];
}

export interface ScanResult {
  sheetsScanned: number;
  scanTruncated: boolean;
}

export const cellAddress = (row: number, col: number) => `${numberToColumn(col + 1)}${row + 1}`;

/**
 * Visit the used range of one sheet (or every visible sheet) in bounded row chunks, optionally
 * limited to `address`. Stops at the executor's scan cap and reports that it did.
 */
export async function scanCells(
  context: Excel.RequestContext,
  target: { sheet?: SheetRef; address?: string },
  visit: (chunk: ScanChunk) => void,
): Promise<ScanResult> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  await context.sync();

  const targets = target.sheet
    ? [resolveSheet(context, target.sheet)]
    : sheets.items.filter((sheet) => sheet.visibility === "Visible");
  if (target.sheet) targets[0]!.load("name");

  let scanned = 0;
  let sheetsScanned = 0;
  let scanTruncated = false;

  outer: for (const sheet of targets) {
    let used = sheet.getUsedRangeOrNullObject(true);
    used.load("isNullObject");
    await context.sync();
    if (used.isNullObject) continue;
    if (target.address) used = used.getIntersectionOrNullObject(sheet.getRange(target.address));
    used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();
    if (used.isNullObject) continue;
    sheetsScanned++;

    const rowsPerChunk = Math.max(1, Math.floor(EXECUTOR_CAPS.readCells / used.columnCount));
    for (let offset = 0; offset < used.rowCount; offset += rowsPerChunk) {
      if (scanned >= EXECUTOR_CAPS.searchScanCells) {
        scanTruncated = true;
        break outer;
      }
      const rows = Math.min(rowsPerChunk, used.rowCount - offset);
      const chunk = sheet.getRangeByIndexes(used.rowIndex + offset, used.columnIndex, rows, used.columnCount);
      chunk.load("values,formulas");
      await context.sync();
      scanned += rows * used.columnCount;
      visit({
        sheet: sheet.name,
        rowIndex: used.rowIndex + offset,
        columnIndex: used.columnIndex,
        rows,
        columns: used.columnCount,
        values: chunk.values,
        formulas: chunk.formulas,
      });
    }
  }
  return { sheetsScanned, scanTruncated };
}

/* -------------------------------------------------------------------------- */
/* find_errors                                                                 */
/* -------------------------------------------------------------------------- */

interface ErrorCell {
  sheet: string;
  address: string;
  error: string;
  formula: string | null;
}

export async function findErrors(input: FindErrorsInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const errors: ErrorCell[] = [];
    const byType: Record<string, number> = {};
    let total = 0;

    const { sheetsScanned, scanTruncated } = await scanCells(context, input, (chunk) => {
      for (let r = 0; r < chunk.rows; r++) {
        for (let c = 0; c < chunk.columns; c++) {
          const value = chunk.values[r]?.[c];
          if (!isErrorValue(value)) continue;
          total++;
          byType[value] = (byType[value] ?? 0) + 1;
          if (errors.length < input.limit) {
            const formula = String(chunk.formulas[r]?.[c] ?? "");
            errors.push({
              sheet: chunk.sheet,
              address: cellAddress(chunk.rowIndex + r, chunk.columnIndex + c),
              error: value,
              formula: formula.startsWith("=") ? formula : null,
            });
          }
        }
      }
    });

    return {
      errors,
      totalFound: total,
      returned: errors.length,
      hasMore: total > errors.length,
      byType,
      sheetsScanned,
      scanTruncated,
      note: total === 0 && !scanTruncated ? "No error values found." : undefined,
    };
  });
}

/* -------------------------------------------------------------------------- */
/* check_formula_consistency                                                   */
/* -------------------------------------------------------------------------- */

function columnNumber(letters: string): number {
  return [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

// No lookbehind: older Office WebViews reject that syntax and would break the whole task pane.
const CELL_REF = /(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])/g;

/**
 * The formula with every cell reference rewritten relative to the cell holding it
 * (`=B2*C2` in D2 and `=B3*C3` in D3 both become `=R[0]C[-2]*R[0]C[-1]`), so formulas that
 * "do the same thing" compare equal. Text literals are left untouched.
 */
export function relativePattern(formula: string, row: number, col: number): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, index) => {
      if (index % 2 === 1) return part;
      return part.replace(CELL_REF, (_m, before: string, colAbs: string, letters: string, rowAbs: string, digits: string) => {
        const c = columnNumber(letters) - 1;
        const r = Number(digits) - 1;
        const rowPart = rowAbs ? `R${r + 1}` : `R[${r - row}]`;
        const colPart = colAbs ? `C${c + 1}` : `C[${c - col}]`;
        return before + rowPart + colPart;
      });
    })
    .join("");
}

const SUMMARY_FORMULA = /^=\s*(SUM|SUBTOTAL|AVERAGE|COUNT|COUNTA|MIN|MAX|AGGREGATE)\s*\(/i;

interface LineCell {
  row: number;
  col: number;
  formula: string | null;
  number: number | null;
}

interface ConsistencyIssue {
  sheet: string;
  address: string;
  kind: "different_formula" | "hardcoded_value";
  formula?: string;
  value?: unknown;
  expected: { address: string; formula: string };
}

/** A pattern must cover at least this share of a line's formulas to count as "the" pattern. */
const DOMINANT_SHARE = 0.6;
const MIN_FORMULAS = 3;

export async function checkFormulaConsistency(input: CheckFormulaConsistencyInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const lines = new Map<string, { sheet: string; cells: LineCell[] }>();

    const { sheetsScanned, scanTruncated } = await scanCells(context, input, (chunk) => {
      for (let r = 0; r < chunk.rows; r++) {
        for (let c = 0; c < chunk.columns; c++) {
          const row = chunk.rowIndex + r;
          const col = chunk.columnIndex + c;
          const raw = String(chunk.formulas[r]?.[c] ?? "");
          const value = chunk.values[r]?.[c];
          const isFormula = raw.startsWith("=");
          const isNumber = !isFormula && typeof value === "number";
          if (!isFormula && !isNumber) continue;
          const key = `${chunk.sheet}\u0000${input.direction === "columns" ? col : row}`;
          let line = lines.get(key);
          if (!line) lines.set(key, (line = { sheet: chunk.sheet, cells: [] }));
          line.cells.push({ row, col, formula: isFormula ? raw : null, number: isNumber ? (value as number) : null });
        }
      }
    });

    const issues: ConsistencyIssue[] = [];
    let linesChecked = 0;

    for (const { sheet, cells } of lines.values()) {
      const formulaCells = cells.filter((cell) => cell.formula !== null);
      if (formulaCells.length < MIN_FORMULAS) continue;

      const groups = new Map<string, LineCell[]>();
      for (const cell of formulaCells) {
        const pattern = relativePattern(cell.formula!, cell.row, cell.col);
        const group = groups.get(pattern);
        if (group) group.push(cell);
        else groups.set(pattern, [cell]);
      }
      const dominant = [...groups.values()].sort((a, b) => b.length - a.length)[0]!;
      if (dominant.length / formulaCells.length < DOMINANT_SHARE) continue;
      linesChecked++;

      const example = dominant[0]!;
      const expected = { address: cellAddress(example.row, example.col), formula: example.formula! };
      const ordered = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);
      const first = ordered.findIndex((cell) => cell.formula !== null);
      const last = ordered.length - 1 - [...ordered].reverse().findIndex((cell) => cell.formula !== null);

      ordered.forEach((cell, index) => {
        const address = cellAddress(cell.row, cell.col);
        if (cell.formula !== null) {
          if (dominant.includes(cell)) return;
          const atEdge = index === first || index === last;
          if (atEdge && SUMMARY_FORMULA.test(cell.formula)) return; // a total at either end is normal
          issues.push({ sheet, address, kind: "different_formula", formula: cell.formula, expected });
        } else if (index > first && index < last) {
          issues.push({ sheet, address, kind: "hardcoded_value", value: cell.number, expected });
        }
      });
    }

    const byKind: Record<string, number> = {};
    for (const issue of issues) byKind[issue.kind] = (byKind[issue.kind] ?? 0) + 1;

    return {
      issues: issues.slice(0, input.limit),
      totalIssues: issues.length,
      returned: Math.min(issues.length, input.limit),
      hasMore: issues.length > input.limit,
      byKind,
      direction: input.direction,
      linesChecked,
      sheetsScanned,
      scanTruncated,
      note:
        issues.length === 0 && !scanTruncated
          ? `No inconsistencies found in ${linesChecked} ${input.direction === "columns" ? "column" : "row"}(s) with a repeated formula pattern.`
          : "Review each item: a different formula or constant can be intentional (totals, overrides, inputs).",
    };
  });
}
