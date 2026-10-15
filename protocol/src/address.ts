export interface CellRef {
  col: number;
  row: number;
}

export interface ParsedRange {
  start: CellRef;
  end: CellRef;
}

export function columnToNumber(letters: string): number {
  let result = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0) - 64;
    if (code < 1 || code > 26) throw new Error(`invalid column letter: ${ch}`);
    result = result * 26 + code;
  }
  return result;
}

export function numberToColumn(value: number): string {
  let n = value;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

/** Parse a single cell like `B7` or `$B$7`. Returns null when malformed. */
export function parseCell(cell: string): CellRef | null {
  const match = CELL_RE.exec(cell.trim());
  if (!match) return null;
  return { col: columnToNumber(match[1]!), row: Number.parseInt(match[2]!, 10) };
}

/** Parse an A1 range like `A1`, `A1:C3`, `$A$1:$C$3`. */
export function parseRange(address: string): ParsedRange | null {
  const parts = address.trim().split(":");
  if (parts.length === 1) {
    const cell = parseCell(parts[0]!);
    return cell ? { start: cell, end: cell } : null;
  }
  if (parts.length !== 2) return null;
  const start = parseCell(parts[0]!);
  const end = parseCell(parts[1]!);
  if (!start || !end) return null;
  return { start, end };
}

export function rangeCellCount(address: string): number | null {
  const parsed = parseRange(address);
  if (!parsed) return null;
  const rows = Math.abs(parsed.end.row - parsed.start.row) + 1;
  const cols = Math.abs(parsed.end.col - parsed.start.col) + 1;
  return rows * cols;
}
