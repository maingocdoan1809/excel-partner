/// <reference lib="dom" />
import { isCellAddress, isOrderedRange, isWholeLine, workbookStem, type CellTarget } from "../../../protocol/src/index.js";

/**
 * Cell and range mentions in the assistant's replies become links that select them in Excel.
 *
 * Two kinds, and only the first is certain:
 *  - A markdown link the assistant wrote with its target spelled out, `[A1:E8](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8)`
 *    (see protocol/src/cell-link.ts). It goes exactly where it says; markdown.ts turns it into a badge with `cellBadge`.
 *  - A reference written in plain words ("F10:F16", "Data!B2", "'Test Cases'!A1:D9"), found here. Only what is unmistakably
 *    a reference becomes a link: plain words that merely look like one (Q4, H2, MP3) are left alone unless they carry a
 *    sheet name or sit in `code`. It names no file, so it goes to a sheet of the workbook the chat is in: never guessed
 *    from the sentence around it, and not offered at all when the reply is about another file (see `mentionsFile`).
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

/** Quarters, half-years and fiscal years read as words, not cells. */
const LOOKS_LIKE_A_WORD = /^(?:Q[1-4]|H[12]|FY\d{2,4})$/;

/** Every reference in `text`, in order. In `code` a lone cell such as Q4 counts, because code is rarely prose. */
export function findCellRefs(text: string, context: "text" | "code" = "text"): CellRef[] {
  const found: CellRef[] = [];
  for (const match of text.matchAll(new RegExp(PATTERN, "gu"))) {
    const [whole, rawSheet, rawRef] = match as unknown as [string, string | undefined, string];
    const sheet = rawSheet?.startsWith("'") ? rawSheet.slice(1, -1).replace(/''/g, "'") : rawSheet;
    const parts = rawRef.split(":");
    const isCells = parts.every(isCellAddress);
    if (!isCells) {
      if (!sheet || !isWholeLine(rawRef)) continue; // whole rows/columns need a sheet
    } else if (parts.length === 2 && !isOrderedRange(parts[0]!, parts[1]!)) {
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

/**
 * True when `text` names one of the workbooks in `names` (a whole word, any case, with or without the extension). A reply
 * that talks about another file cannot have its bare cell references linked: they could be that file's, and a link into
 * the wrong file is worse than no link. Used only to hold links back, never to decide where one goes.
 */
export function mentionsFile(text: string, names: readonly string[]): boolean {
  return mentionsName(text, names.map(workbookStem));
}

/**
 * True when `text` names one of the sheets in `names` (a whole word, any case). A bare "C5" in a reply that also says
 * "sheet2/C5" is not the open sheet's: it could be the named sheet's, and the page cannot tell which, so it is left unlinked.
 */
export function mentionsSheet(text: string, names: readonly string[]): boolean {
  return mentionsName(text, names);
}

function mentionsName(text: string, names: readonly string[]): boolean {
  for (const name of names) {
    // One letter would match the article "a" in a sentence; such a name is too short to tell from ordinary words.
    if (name.length < 2) continue;
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp(String.raw`(?<![\p{L}\p{N}_])${escaped}(?![\p{L}\p{N}_])`, "iu").test(text)) return true;
  }
  return false;
}

/** The badge for a reference: it carries where it goes, and the click handler in main.ts does the rest. */
export function cellBadge(text: string, target: CellTarget): HTMLAnchorElement {
  const link = document.createElement("a");
  link.className = "cell-link";
  link.href = "#";
  link.textContent = text;
  link.dataset.address = target.address;
  if (target.file) link.dataset.file = target.file;
  if (target.sheet) link.dataset.sheet = target.sheet;
  const where = `${target.file ? `${target.file} → ` : ""}${target.sheet ? `${target.sheet}!` : ""}${target.address}`;
  link.title = `Select ${where} in Excel${target.sheet || target.file ? "" : " (on the sheet that is open)"}`;
  return link;
}

const SKIP = new Set(["A", "PRE", "SCRIPT", "STYLE", "BUTTON"]);

/**
 * Replace plain-word references inside `root` with links. A reference without a sheet name goes to the last sheet named
 * earlier in the same text, or else to whichever sheet is open when it is clicked.
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
      const sheet = ref.sheet ?? current;
      pieces.push(node.data.slice(at, ref.start), cellBadge(ref.text, { ...(sheet ? { sheet } : {}), address: ref.address }));
      at = ref.end;
    }
    pieces.push(node.data.slice(at));
    (inCode ? parent : node).replaceWith(...pieces.filter((piece) => piece !== ""));
  }
}
