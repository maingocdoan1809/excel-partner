/// <reference lib="dom" />
import { columnToNumber } from "../../../protocol/src/index.js";

/**
 * Turns cell and range mentions in the assistant's text ("F10:F16", "Data!B2", "'Test Cases'!A1:D9") into links
 * that select them in Excel. Only what is unmistakably a reference becomes a link: plain words that merely look like
 * one (Q4, H2, MP3) are left alone unless they carry a sheet name or sit in `code`.
 */

export interface CellRef {
  /** Sheet named in the text; undefined when the text names none. */
  sheet?: string;
  /** A1 address without `$`, e.g. `F10:F16`. */
  address: string;
  /** The text as written. */
  text: string;
  start: number;
  end: number;
}

const CELL = String.raw`\$?[A-Z]{1,3}\$?\d{1,7}`;
const SHEET = String.raw`(?:'(?:[^']|'')+'|[\p{L}\p{N}_.]+)`;
// Whole columns/rows (D:D, 3:5) look like ratios and times, so they only count with a sheet name in front.
const WHOLE = String.raw`\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d{1,7}:\$?\d{1,7}`;
// Not preceded by a word character or `#` (ids like ID#A1) and not followed by one, by `(` (LOG10(), or by `!` that
// leads into another reference (a sheet named A1) -- a closing "G1!" at the end of a sentence still counts.
const PATTERN = String.raw`(?<![\p{L}\p{N}_$.'!:#])(?:(${SHEET})!)?(${CELL}(?::${CELL})?|${WHOLE})(?![\p{L}\p{N}_(:]|!(?=[\p{L}\p{N}_'$]))`;

const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;
/** Quarters, half-years and fiscal years read as words, not cells. */
const LOOKS_LIKE_A_WORD = /^(?:Q[1-4]|H[12]|FY\d{2,4})$/;

function validCell(cell: string): boolean {
  const match = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/.exec(cell);
  if (!match) return false;
  return columnToNumber(match[1]!) <= MAX_COLUMN && Number(match[2]) >= 1 && Number(match[2]) <= MAX_ROW;
}

function validWhole(part: string): boolean {
  const [from = "", to = ""] = part.replace(/\$/g, "").split(":");
  if (/^\d+$/.test(from)) return Number(from) >= 1 && Number(to) <= MAX_ROW && Number(from) <= Number(to);
  return columnToNumber(from) <= MAX_COLUMN && columnToNumber(to) <= MAX_COLUMN && columnToNumber(from) <= columnToNumber(to);
}

/** Every reference in `text`, in order. In `code` a lone cell such as Q4 counts, because code is rarely prose. */
export function findCellRefs(text: string, context: "text" | "code" = "text"): CellRef[] {
  const found: CellRef[] = [];
  for (const match of text.matchAll(new RegExp(PATTERN, "gu"))) {
    const [whole, rawSheet, rawRef] = match as unknown as [string, string | undefined, string];
    const sheet = rawSheet?.startsWith("'") ? rawSheet.slice(1, -1).replace(/''/g, "'") : rawSheet;
    const parts = rawRef.split(":");
    const isCells = parts.every(validCell);
    if (!isCells) {
      if (!sheet || !validWhole(rawRef)) continue; // whole rows/columns need a sheet
    } else if (parts.length === 2 && !validRange(parts as [string, string])) {
      continue;
    }
    // A lone cell in running text could be a word (Q4); with a sheet or in code it is a cell.
    if (isCells && parts.length === 1 && !sheet && context === "text" && LOOKS_LIKE_A_WORD.test(rawRef.replace(/\$/g, ""))) continue;
    found.push({ ...(sheet ? { sheet } : {}), address: rawRef.replace(/\$/g, ""), text: whole, start: match.index!, end: match.index! + whole.length });
  }
  return found;
}

/** True when `text` is nothing but references and separators (`C10, C14 and C16`), never a formula like =SUM(C1:C5). */
export function isRefList(text: string, refs: CellRef[]): boolean {
  if (refs.length === 0) return false;
  let rest = "";
  let at = 0;
  for (const ref of refs) {
    rest += text.slice(at, ref.start);
    at = ref.end;
  }
  rest += text.slice(at);
  return /^[\s,;]*$/.test(rest.replace(/(?<![\p{L}\p{N}_])(?:và|and|or|hoặc)(?![\p{L}\p{N}_])/giu, ""));
}

function validRange([from, to]: [string, string]): boolean {
  const a = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(from)!;
  const b = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(to)!;
  return columnToNumber(a[1]!) <= columnToNumber(b[1]!) && Number(a[2]) <= Number(b[2]);
}

function makeLink(ref: CellRef, sheet: string | undefined): HTMLAnchorElement {
  const link = document.createElement("a");
  link.className = "cell-link";
  link.href = "#";
  link.textContent = ref.text;
  link.dataset.address = ref.address;
  if (sheet) link.dataset.sheet = sheet;
  link.title = `Select ${sheet ? `${sheet}!` : ""}${ref.address} in Excel${sheet ? "" : " (on the sheet that is open)"}`;
  return link;
}

const SKIP = new Set(["A", "PRE", "SCRIPT", "STYLE", "BUTTON"]);

/**
 * Replace references inside `root` with links. A reference without a sheet name goes to the last sheet named earlier
 * in the same text, or else to whichever sheet is open when it is clicked.
 */
export function linkifyCells(root: HTMLElement): void {
  let current: string | undefined;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const texts: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) texts.push(node as Text);
  for (const node of texts) {
    const parent = node.parentElement;
    if (!parent) continue;
    let skipped = false;
    for (let el: HTMLElement | null = parent; el && el !== root; el = el.parentElement) if (SKIP.has(el.tagName)) { skipped = true; break; }
    if (skipped) continue;
    const inCode = parent.tagName === "CODE";
    const refs = findCellRefs(node.data, inCode ? "code" : "text");
    if (refs.length === 0) continue;
    // Inline code links only when it is nothing but references, so formulas like =SUM(F10:F16) stay text. The code box
    // is replaced by the badges: a badge inside a code box would draw two bubbles.
    if (inCode && !(parent.childNodes.length === 1 && isRefList(node.data, refs))) continue;
    const pieces: Array<string | HTMLElement> = [];
    let at = 0;
    for (const ref of refs) {
      if (ref.sheet) current = ref.sheet;
      pieces.push(node.data.slice(at, ref.start), makeLink(ref, ref.sheet ?? current));
      at = ref.end;
    }
    pieces.push(node.data.slice(at));
    (inCode ? parent : node).replaceWith(...pieces.filter((piece) => piece !== ""));
  }
}
