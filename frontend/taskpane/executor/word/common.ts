import { ToolFailure } from "../common.js";
import { requireSupport, type Requirement } from "../../host/requirements.js";

/** Word tool results are cut to this many JSON characters before they reach the model. */
export const RESULT_LIMIT = 30_000;

/** Requirement sets the Word executor needs beyond the ones `host/requirements.ts` names. */
export const WORD_API_1_3: Requirement = ["WordApi", "1.3"];
export const WORD_TRACK_MODE: Requirement = ["WordApi", "1.4"];
export const WORD_PAGE_SETUP: Requirement = ["WordApiDesktop", "1.3"];

export { requireSupport };

/** Run `fn` inside `Word.run`, or stop with `unsupported` when this pane has no Word API (browser preview, Excel, PowerPoint). */
export async function wordRun<T>(fn: (context: Word.RequestContext) => Promise<T>): Promise<T> {
  if (typeof Word === "undefined" || typeof Word.run !== "function") {
    throw new ToolFailure("unsupported", "The Word API is not available in this task pane. Open the add-in inside Word.");
  }
  return Word.run(fn);
}

/** Office reports a missing style / bookmark / item as `ItemNotFound`; for Word that is not a sheet problem. */
export function asWordFailure(error: unknown): unknown {
  const e = error as { code?: string; message?: string };
  if (!(error instanceof ToolFailure) && e?.code === "ItemNotFound") {
    return new ToolFailure("range_not_found", `${e.message ?? "Item not found"} (the style, bookmark, comment or object named in the call does not exist in this document)`, { officeCode: e.code });
  }
  return error;
}

/** `text` cut to `max` characters with a visible marker. */
export function clip(text: string, max: number): { text: string; truncated: boolean } {
  if (text.length <= max) return { text, truncated: false };
  return { text: `${text.slice(0, max)}…`, truncated: true };
}

/** Short one-line preview of document text for summaries. */
export function preview(text: string, max = 60): string {
  const flat = text.replace(/[\r\n\u000b\u0007]+/g, " ⏎ ").trim();
  return flat.length <= max ? flat : `${flat.slice(0, max)}…`;
}

/** Word returns `\r` for paragraph breaks; the model works with `\n`. */
export function normalizeNewlines(text: string): string {
  return text.replace(/\r\n|\r|\u000b/g, "\n");
}

/**
 * Build a result that fits `RESULT_LIMIT`: `build(n)` renders the result with the first `n` list entries. Returns the
 * rendering with the most entries that still fits (flagging `truncated` when it had to drop any).
 */
export function fitResult<T extends Record<string, unknown>>(total: number, build: (count: number) => T, limit = RESULT_LIMIT): T & { truncated?: boolean } {
  const size = (value: unknown) => JSON.stringify(value).length;
  const full = build(total);
  if (size(full) <= limit) return full;
  let low = 0;
  let high = total - 1;
  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    if (size(build(mid)) <= limit) low = mid;
    else high = mid - 1;
  }
  return { ...build(low), truncated: true };
}

/** Final guard for every Word result. */
export function assertWithinLimit(data: unknown): void {
  const size = JSON.stringify(data ?? null).length;
  if (size > RESULT_LIMIT) {
    throw new ToolFailure("size_limit_exceeded", `The result is ${size} characters, over the ${RESULT_LIMIT} limit. Ask for a smaller part of the document.`);
  }
}

export function isBlankText(text: string): boolean {
  return /^[ \t\r\n\u00a0]*$/.test(text);
}

export function assertParagraphIndex(index: number, count: number, what = "paragraph_index"): void {
  if (!Number.isInteger(index) || index < 0 || index >= count) {
    throw new ToolFailure("invalid_arguments", `${what} ${index} is outside the document, which has ${count} paragraph${count === 1 ? "" : "s"} (0..${Math.max(0, count - 1)}). Read the outline again for current indexes.`);
  }
}
