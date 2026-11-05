import { columnToNumber } from "./address.js";

/**
 * A reference to cells that the assistant states in full, as an ordinary markdown link:
 *
 *   [A1:E8](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8)
 *
 * `range` is required; `sheet` is the sheet (the open one when absent); `file` is another open workbook (this one when
 * absent). Spaces in a name are written %20. The task pane reads the link with the markdown and URL parsers, never by
 * guessing from the words around it, so a link goes exactly where it says or nowhere.
 */
export interface CellTarget {
  file?: string;
  sheet?: string;
  /** A1 address without `$`: `B2`, `A1:E8`, `D:D` or `3:5`. */
  address: string;
}

const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;
const MAX_NAME = 255;

/** One cell such as `B2` or `$B$2`. */
export function isCellAddress(cell: string): boolean {
  const match = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/.exec(cell);
  if (!match) return false;
  return columnToNumber(match[1]!) <= MAX_COLUMN && Number(match[2]) >= 1 && Number(match[2]) <= MAX_ROW;
}

/** Whole columns (`D:D`) or rows (`3:5`), in order. */
export function isWholeLine(part: string): boolean {
  const [from = "", to = ""] = part.replace(/\$/g, "").split(":");
  if (/^\d+$/.test(from)) return /^\d+$/.test(to) && Number(from) >= 1 && Number(to) <= MAX_ROW && Number(from) <= Number(to);
  if (!/^[A-Z]{1,3}$/.test(from) || !/^[A-Z]{1,3}$/.test(to)) return false;
  return columnToNumber(from) <= MAX_COLUMN && columnToNumber(to) <= MAX_COLUMN && columnToNumber(from) <= columnToNumber(to);
}

/** Two cells, the first not after the second. */
export function isOrderedRange(from: string, to: string): boolean {
  const a = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(from);
  const b = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(to);
  if (!a || !b) return false;
  return columnToNumber(a[1]!) <= columnToNumber(b[1]!) && Number(a[2]) <= Number(b[2]);
}

/** A cell, a range of cells, or whole rows or columns. */
export function isA1Range(address: string): boolean {
  const parts = address.split(":");
  if (parts.length === 1) return isCellAddress(address);
  if (parts.length !== 2) return false;
  if (parts.every(isCellAddress)) return isOrderedRange(parts[0]!, parts[1]!);
  return isWholeLine(address);
}

/** The `cell:` link address for `target`. */
export function cellHref(target: CellTarget): string {
  const parts: string[] = [];
  if (target.file) parts.push(`file=${encodeURIComponent(target.file)}`);
  if (target.sheet) parts.push(`sheet=${encodeURIComponent(target.sheet)}`);
  parts.push(`range=${target.address.replace(/\$/g, "")}`);
  return `cell:?${parts.join("&")}`;
}

/** A markdown link to `target` with `label` as its text. */
export function cellLinkMarkdown(label: string, target: CellTarget): string {
  return `[${label.replace(/[[\]\\]/g, "\\$&")}](${cellHref(target)})`;
}

/** What a `cell:` link points at, or null when it is not one or does not name a real range. */
export function parseCellHref(href: string): CellTarget | null {
  let url: URL;
  try { url = new URL(href.trim()); } catch { return null; }
  if (url.protocol !== "cell:") return null;
  const address = (url.searchParams.get("range") ?? "").replace(/\$/g, "").trim().toUpperCase();
  if (!isA1Range(address)) return null;
  const file = (url.searchParams.get("file") ?? "").trim();
  const sheet = (url.searchParams.get("sheet") ?? "").trim();
  if (file.length > MAX_NAME || sheet.length > MAX_NAME) return null;
  return { ...(file ? { file } : {}), ...(sheet ? { sheet } : {}), address };
}
