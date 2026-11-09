import { ToolFailure } from "../common.js";
import { normalizeNewlines, preview } from "./common.js";

/** Word's `search` accepts at most 255 characters; chunks stay a little under so special codes never tip it over. */
export const SEARCH_LIMIT = 255;
const CHUNK = 250;
/** Start x end candidates tried for a long text before giving up as too ambiguous. */
const MAX_PAIRS = 40;

interface Token { raw: string; escaped: string }

/** Split text into units whose escaped form is what Word's search must see: `^` -> `^^`, a line break -> `^p`, a tab -> `^t`. */
function tokenize(text: string): Token[] {
  const tokens: Token[] = [];
  const chars = Array.from(text);
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i]!;
    if (ch === "\r" && chars[i + 1] === "\n") { tokens.push({ raw: "\r\n", escaped: "^p" }); i++; }
    else if (ch === "\n" || ch === "\r") tokens.push({ raw: ch, escaped: "^p" });
    else if (ch === "\t") tokens.push({ raw: ch, escaped: "^t" });
    else if (ch === "^") tokens.push({ raw: ch, escaped: "^^" });
    else tokens.push({ raw: ch, escaped: ch });
  }
  return tokens;
}

/** Text as Word's search wants it (wildcards are always off, so only `^` is special). */
export function escapeSearchText(text: string): string {
  return tokenize(text).map((t) => t.escaped).join("");
}

export type SearchPlan = { kind: "short"; term: string } | { kind: "long"; head: string; tail: string };

export function planSearch(text: string): SearchPlan {
  const tokens = tokenize(text);
  const total = tokens.reduce((sum, t) => sum + t.escaped.length, 0);
  if (total <= SEARCH_LIMIT) return { kind: "short", term: tokens.map((t) => t.escaped).join("") };
  let head = "";
  for (const t of tokens) { if (head.length + t.escaped.length > CHUNK) break; head += t.escaped; }
  let tail = "";
  for (let i = tokens.length - 1; i >= 0; i--) { const t = tokens[i]!; if (tail.length + t.escaped.length > CHUNK) break; tail = t.escaped + tail; }
  return { kind: "long", head, tail };
}

export interface LocateOptions {
  matchCase: boolean;
  matchWholeWord?: boolean;
}

export interface Located {
  /** Every match in document order. */
  ranges: Word.Range[];
  /** How the matches were found. */
  method: "search" | "start_end";
}

function compareKey(text: string, matchCase: boolean): string {
  const base = normalizeNewlines(text).replace(/\u0007/g, "").normalize("NFC");
  return matchCase ? base : base.toLowerCase();
}

/**
 * Every place `text` occurs in `scope`. Text up to 255 characters is one `search` call. Longer text is found through its
 * first and last 250 characters: each start x end pair is spanned with `expandTo` and kept only when the spanned text is
 * exactly the wanted text, so a long text is never guessed from its ends alone.
 */
export async function locateText(context: Word.RequestContext, scope: Word.Body | Word.Range, text: string, options: LocateOptions): Promise<Located> {
  const searchOptions = { matchCase: options.matchCase, matchWholeWord: options.matchWholeWord ?? false, matchWildcards: false };
  const plan = planSearch(text);
  if (plan.kind === "short") {
    const hits = scope.search(plan.term, searchOptions);
    hits.load("items");
    await context.sync();
    return { ranges: [...hits.items], method: "search" };
  }
  const starts = scope.search(plan.head, searchOptions);
  const ends = scope.search(plan.tail, searchOptions);
  starts.load("items");
  ends.load("items");
  await context.sync();
  if (starts.items.length === 0 || ends.items.length === 0) return { ranges: [], method: "start_end" };
  if (starts.items.length * ends.items.length > MAX_PAIRS) {
    throw new ToolFailure("invalid_arguments", `The beginning of old_text appears ${starts.items.length} times and its end ${ends.items.length} times, too many to tell which span is meant. Use a shorter old_text, or one that starts or ends with a more distinctive phrase.`);
  }
  const spans: Array<{ range: Word.Range; order: number }> = [];
  starts.items.forEach((start, i) => ends.items.forEach((end, j) => {
    const span = start.expandTo(end);
    span.load("text");
    spans.push({ range: span, order: i * 1000 + j });
  }));
  await context.sync();
  const wanted = compareKey(text, options.matchCase);
  const matching = spans.filter((s) => compareKey(s.range.text, options.matchCase) === wanted);
  matching.sort((a, b) => a.order - b.order);
  return { ranges: matching.map((s) => s.range), method: "start_end" };
}

export interface PickedTarget {
  ranges: Word.Range[];
  total: number;
  method: Located["method"];
}

/** Choose the range(s) an edit applies to: every match, or the `occurrence`-th (1-based). Fails with a hint when there is none. */
export async function pickTargets(
  context: Word.RequestContext,
  scope: Word.Body | Word.Range,
  text: string,
  options: LocateOptions & { occurrence: number; all?: boolean; what?: string },
): Promise<PickedTarget> {
  const located = await locateText(context, scope, text, options);
  const what = options.what ?? "text";
  if (located.ranges.length === 0) {
    throw new ToolFailure("range_not_found", `The ${what} “${preview(text, 80)}” does not appear in the document${options.matchCase ? " (the match is case-sensitive; check capitalisation, or search_doc_text for the exact wording)" : ""}. It may span a different formatting or wording than you expect; use search_doc_text or read_doc_section to see the current text.`, { matches: 0 });
  }
  if (options.all) return { ranges: located.ranges, total: located.ranges.length, method: located.method };
  if (options.occurrence > located.ranges.length) {
    throw new ToolFailure("range_not_found", `Occurrence ${options.occurrence} of the ${what} “${preview(text, 80)}” does not exist: it appears ${located.ranges.length} time${located.ranges.length === 1 ? "" : "s"}.`, { matches: located.ranges.length });
  }
  return { ranges: [located.ranges[options.occurrence - 1]!], total: located.ranges.length, method: located.method };
}
