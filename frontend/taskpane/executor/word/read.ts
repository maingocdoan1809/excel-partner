import type { GetDocOutlineInput, ReadDocSectionInput, SearchDocTextInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { assertParagraphIndex, clip, fitResult, normalizeNewlines, preview, wordRun } from "./common.js";
import { loadParagraphs } from "./paragraphs.js";
import { escapeSearchText } from "./search.js";

const PARAGRAPH_TEXT_LIMIT = 5000;

export async function getDocOutline(input: GetDocOutlineInput): Promise<unknown> {
  return wordRun(async (context) => {
    const { info } = await loadParagraphs(context, { structure: false });
    const headings = info.filter((p) => p.headingLevel !== null);
    const shown = headings.slice(0, input.max_headings ?? 300);
    return fitResult(shown.length, (n) => ({
      paragraphCount: info.length,
      headingCount: headings.length,
      headings: shown.slice(0, n).map((p) => ({ index: p.index, level: p.headingLevel, text: clip(normalizeNewlines(p.text).trim(), 200).text, style: p.style })),
      truncated: n < headings.length,
    }));
  });
}

export async function readDocSection(input: ReadDocSectionInput): Promise<unknown> {
  return wordRun(async (context) => {
    const { info } = await loadParagraphs(context);
    if (info.length === 0) return { paragraphCount: 0, startIndex: 0, endIndex: -1, paragraphs: [], truncated: false, summary: "The document has no paragraphs." };
    const count = input.count ?? 40;
    let start = 0;
    let end = info.length; // exclusive
    let heading: { index: number; level: number; text: string } | undefined;
    if (input.paragraph_index !== undefined) assertParagraphIndex(input.paragraph_index, info.length);
    if (input.heading !== undefined) {
      const wanted = input.heading.trim().toLowerCase();
      const from = input.paragraph_index ?? 0;
      const found = info.find((p) => p.index >= from && p.headingLevel !== null && normalizeNewlines(p.text).trim().toLowerCase() === wanted);
      if (!found) {
        const known = info.filter((p) => p.headingLevel !== null).slice(0, 25).map((p) => `“${preview(p.text, 50)}”`);
        throw new ToolFailure("range_not_found", `No heading “${preview(input.heading, 80)}” was found${from > 0 ? ` from paragraph ${from}` : ""}.${known.length ? ` Headings in the document: ${known.join(", ")}.` : " The document has no headings; read by paragraph_index instead."}`);
      }
      heading = { index: found.index, level: found.headingLevel!, text: preview(found.text, 200) };
      start = found.index;
      const next = info.find((p) => p.index > found.index && p.headingLevel !== null && p.headingLevel <= found.headingLevel!);
      end = next ? next.index : info.length;
    } else if (input.paragraph_index !== undefined) {
      start = input.paragraph_index;
    }
    const window = info.slice(start, Math.min(end, start + count));
    const sectionEnd = end;
    const rows = window.map((p) => {
      const text = clip(normalizeNewlines(p.text), PARAGRAPH_TEXT_LIMIT);
      return {
        index: p.index,
        style: p.style,
        ...(p.headingLevel !== null ? { heading_level: p.headingLevel } : {}),
        ...(p.listItem ? { list_item: true } : {}),
        ...(p.inTable ? { in_table: true } : {}),
        text: text.text,
        ...(text.truncated ? { text_truncated: true } : {}),
      };
    });
    return fitResult(rows.length, (n) => {
      const shown = rows.slice(0, n);
      const nextIndex = start + shown.length;
      return {
        paragraphCount: info.length,
        ...(heading ? { heading } : {}),
        startIndex: start,
        endIndex: start + shown.length - 1,
        paragraphs: shown,
        ...(nextIndex < sectionEnd ? { nextIndex } : {}),
        truncated: nextIndex < sectionEnd,
      };
    });
  });
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export async function searchDocText(input: SearchDocTextInput): Promise<unknown> {
  return wordRun(async (context) => {
    const matchCase = input.match_case ?? false;
    const wholeWord = input.whole_word ?? false;
    const maxResults = input.max_results ?? 20;
    // Word's own count is the one edit_doc_text's `occurrence` follows, so ask Word for it...
    const hits = context.document.body.search(escapeSearchText(input.query), { matchCase, matchWholeWord: wholeWord, matchWildcards: false });
    hits.load("items");
    const { info } = await loadParagraphs(context, { structure: false });
    const total = hits.items.length;
    // ...and read positions from the paragraph text, because a Word range does not say which paragraph it is in.
    const pattern = escapeRegExp(normalizeNewlines(input.query));
    const regex = new RegExp(wholeWord ? `(?<![\\p{L}\\p{N}_])${pattern}(?![\\p{L}\\p{N}_])` : pattern, matchCase ? "gu" : "giu");
    const matches: Array<{ occurrence: number; paragraph_index: number; style: string; context: string }> = [];
    let found = 0;
    for (const p of info) {
      const text = normalizeNewlines(p.text);
      for (const m of text.matchAll(regex)) {
        found++;
        if (matches.length < maxResults) {
          const at = m.index ?? 0;
          const from = Math.max(0, at - 60);
          const snippet = text.slice(from, at + m[0].length + 60).replace(/\n/g, " ⏎ ");
          matches.push({ occurrence: found, paragraph_index: p.index, style: p.style, context: `${from > 0 ? "…" : ""}${snippet}${at + m[0].length + 60 < text.length ? "…" : ""}` });
        }
      }
    }
    const note = total !== found
      ? `Word counted ${total} match${total === 1 ? "" : "es"} but only ${found} could be placed inside single paragraphs (the rest span paragraphs, fields or special characters); occurrence numbers follow Word's order.`
      : undefined;
    return fitResult(matches.length, (n) => ({
      query: input.query,
      total,
      returned: Math.min(n, matches.length),
      matches: matches.slice(0, n),
      truncated: total > Math.min(n, matches.length),
      ...(note ? { note } : {}),
    }));
  });
}
