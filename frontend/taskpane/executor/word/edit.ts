import type { CollapseBlankParagraphsInput, EditDocListInput, EditDocTextInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { isSupported, WORD_LISTS, WORD_TRACK } from "../../host/requirements.js";
import { assertParagraphIndex, clip, isBlankText, preview, requireSupport, WORD_API_1_3, wordRun } from "./common.js";
import { deleteParagraph } from "./paragraphs.js";
import { pickTargets, type PickedTarget } from "./search.js";

/* -------------------------------------------------------------------------- */
/* edit_doc_text                                                               */
/* -------------------------------------------------------------------------- */

export interface ReplaceSpec {
  oldText: string;
  newText: string;
  matchCase: boolean;
  occurrence: number;
  all: boolean;
}

/** Queue the replacement of every target range. An empty replacement deletes the match. */
export function replaceRanges(targets: PickedTarget, newText: string): void {
  for (const range of targets.ranges) {
    if (newText === "") range.delete();
    else range.insertText(newText, "Replace");
  }
}

/**
 * Replace text found by one search. The same path serves edit_doc_text and propose_doc_edits: a text edit is always
 * `insertText(..., "Replace")` on the hit, never a rebuild of the body.
 */
export async function replaceText(context: Word.RequestContext, spec: ReplaceSpec): Promise<PickedTarget> {
  const targets = await pickTargets(context, context.document.body, spec.oldText, {
    matchCase: spec.matchCase,
    occurrence: spec.occurrence,
    all: spec.all,
    what: "old_text",
  });
  replaceRanges(targets, spec.newText);
  await context.sync();
  return targets;
}

export async function editDocText(input: EditDocTextInput): Promise<unknown> {
  const track = input.track === true;
  if (track) {
    requireSupport("Editing with Track Changes on", WORD_TRACK, "Edit without track, or ask the person to turn Track Changes on in the Review tab and then edit.");
  }
  return wordRun(async (context) => {
    const spec: ReplaceSpec = {
      oldText: input.old_text,
      newText: input.new_text,
      matchCase: input.match_case ?? true,
      occurrence: input.occurrence ?? 1,
      all: input.all ?? false,
    };
    const targets = await pickTargets(context, context.document.body, spec.oldText, { matchCase: spec.matchCase, occurrence: spec.occurrence, all: spec.all, what: "old_text" });
    let turnedOn = false;
    if (track) {
      context.document.load("changeTrackingMode");
      await context.sync();
      if (context.document.changeTrackingMode === "Off") turnedOn = true;
      context.document.changeTrackingMode = "TrackAll";
    }
    replaceRanges(targets, spec.newText);
    await context.sync();
    const replaced = targets.ranges.length;
    const others = spec.all ? 0 : targets.total - 1;
    return {
      summary: `${spec.newText === "" ? "Deleted" : "Replaced"} ${replaced} match${replaced === 1 ? "" : "es"}: “${preview(spec.oldText, 40)}”${spec.newText === "" ? "" : ` → “${preview(spec.newText, 40)}”`}${track ? " (tracked)" : ""}`,
      replaced,
      matches: targets.total,
      ...(spec.all ? {} : { occurrence: spec.occurrence }),
      ...(others > 0 ? { note: `The text appears ${targets.total} times; only occurrence ${spec.occurrence} was changed. Pass occurrence or all to change others.` } : {}),
      ...(targets.method === "start_end" ? { matched_by: "start_and_end of a long text (verified identical)" } : {}),
      tracked: track ? true : undefined,
      ...(turnedOn ? { track_changes_left_on: true } : {}),
      old_text: clip(spec.oldText, 200).text,
      new_text: clip(spec.newText, 200).text,
    };
  });
}

/* -------------------------------------------------------------------------- */
/* edit_doc_list                                                               */
/* -------------------------------------------------------------------------- */

interface ListEntry { text: string; level: number }

function htmlEscape(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Nested `<ul>/<ol>` for the WordApi 1.1 fallback. */
export function listHtml(tag: "ul" | "ol", entries: ListEntry[]): string {
  let html = "";
  let depth = -1;
  for (const entry of entries) {
    const level = Math.min(entry.level, depth + 1);
    while (depth > level) { html += `</li></${tag}>`; depth--; }
    if (depth === level) html += "</li>";
    while (depth < level) { html += `<${tag}>`; depth++; }
    html += `<li>${htmlEscape(entry.text)}`;
  }
  while (depth >= 0) { html += `</li></${tag}>`; depth--; }
  return html;
}

const BULLETS = ["Solid", "Hollow", "Square"] as const;
const NUMBERINGS = ["Arabic", "LowerLetter", "LowerRoman"] as const;

export async function editDocList(input: EditDocListInput): Promise<unknown> {
  const entries: ListEntry[] = input.items.map((item) => {
    const text = typeof item === "string" ? item : item.text;
    const level = typeof item === "string" ? 0 : (item.level ?? 0);
    return { text: text.replace(/[\r\n\u000b]+/g, " ").trim(), level };
  }).filter((e) => e.text.length > 0);
  if (entries.length === 0) throw new ToolFailure("invalid_arguments", "items has no text to insert.");
  const tag = input.operation === "numbered" ? "ol" : "ul";
  return wordRun(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const count = paragraphs.items.length;
    const after = input.after_paragraph;
    if (after !== undefined) assertParagraphIndex(after, count, "after_paragraph");
    const firstIndex = after !== undefined ? after + 1 : count;
    const common = { items: entries.length, first_paragraph_index: firstIndex, ...(after !== undefined ? { after_paragraph: after } : {}) };

    if (isSupported(WORD_LISTS)) {
      const created: Word.Paragraph[] = [];
      entries.forEach((entry, i) => {
        const previous = created[i - 1];
        const paragraph = i === 0
          ? (after !== undefined ? paragraphs.items[after]!.insertParagraph(entry.text, "After") : body.insertParagraph(entry.text, "End"))
          : previous!.insertParagraph(entry.text, "After");
        paragraph.styleBuiltIn = "ListParagraph";
        created.push(paragraph);
      });
      const list = created[0]!.startNewList();
      list.load("id");
      await context.sync();
      created.forEach((paragraph, i) => {
        if (i > 0) paragraph.attachToList(list.id, entries[i]!.level);
      });
      if (entries[0]!.level > 0) created[0]!.listItem.level = entries[0]!.level;
      await context.sync();
      let note: string | undefined;
      try {
        const levels = [...new Set(entries.map((e) => e.level))];
        for (const level of levels) {
          if (tag === "ul") list.setLevelBullet(level, BULLETS[level % BULLETS.length]!);
          else list.setLevelNumbering(level, NUMBERINGS[level % NUMBERINGS.length]!, [level, "."]);
        }
        await context.sync();
      } catch {
        note = "The list was created, but its bullet / number style could not be set; Word's default list formatting is used.";
      }
      return { summary: `Inserted a ${input.operation} list of ${entries.length} item${entries.length === 1 ? "" : "s"}`, ...common, method: "list_api", ...(note ? { note } : {}) };
    }

    const html = listHtml(tag, entries);
    if (after !== undefined) paragraphs.items[after]!.insertParagraph("", "After").insertHtml(html, "Replace");
    else body.insertHtml(html, "End");
    await context.sync();
    return {
      summary: `Inserted a ${input.operation} list of ${entries.length} item${entries.length === 1 ? "" : "s"} (as HTML; this Word has no List API)`,
      ...common,
      method: "html",
      note: "Inserted as HTML because WordApi 1.3 is not available; an empty paragraph may remain before the list.",
    };
  });
}

/* -------------------------------------------------------------------------- */
/* collapse_blank_paragraphs                                                   */
/* -------------------------------------------------------------------------- */

const DELETE_BATCH = 400;

export async function collapseBlankParagraphs(input: CollapseBlankParagraphsInput): Promise<unknown> {
  requireSupport("collapse_blank_paragraphs", WORD_API_1_3, "It needs table nesting information so it never deletes inside tables.");
  const max = input.max ?? 1;
  const previewOnly = input.preview ?? false;
  return wordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text,items/tableNestingLevel");
    await context.sync();
    const items = paragraphs.items;
    const texts = items.map((p) => p.text);
    // Blank means only spaces, tabs, line ends and non-breaking spaces; a page or section break (\f) is not blank.
    const textBlank = items.map((p, i) => p.tableNestingLevel === 0 && isBlankText(texts[i]!));

    const runsOf = (blank: boolean[]): Array<{ start: number; end: number }> => {
      const runs: Array<{ start: number; end: number }> = [];
      for (let i = 0; i < blank.length; i++) {
        if (!blank[i]) continue;
        let j = i;
        while (j + 1 < blank.length && blank[j + 1]) j++;
        runs.push({ start: i, end: j });
        i = j;
      }
      return runs;
    };

    // A paragraph that only looks empty may hold a picture or a content control: check just the ones in runs that could shrink.
    const candidates = runsOf(textBlank).filter((r) => r.end - r.start + 1 > max).flatMap((r) => Array.from({ length: r.end - r.start + 1 }, (_, k) => r.start + k));
    const checks = candidates.map((i) => {
      const pictures = items[i]!.inlinePictures;
      const controls = items[i]!.contentControls;
      pictures.load("items");
      controls.load("items");
      return { i, pictures, controls };
    });
    if (checks.length > 0) await context.sync();
    const holdsObject = new Set(checks.filter((c) => c.pictures.items.length > 0 || c.controls.items.length > 0).map((c) => c.i));
    const blank = textBlank.map((b, i) => b && !holdsObject.has(i));

    const lastIndex = items.length - 1;
    const toDelete: number[] = [];
    const runs: Array<{ start_index: number; length: number; removed: number }> = [];
    for (const run of runsOf(blank)) {
      const length = run.end - run.start + 1;
      if (length <= max) continue;
      // Keep the first `max` of a run; in the run that ends the document keep the last `max`, because the final paragraph cannot be deleted.
      const first = run.end === lastIndex ? run.start : run.start + max;
      const doomed = Array.from({ length: length - max }, (_, k) => first + k).filter((i) => i !== lastIndex);
      if (doomed.length === 0) continue;
      toDelete.push(...doomed);
      runs.push({ start_index: run.start, length, removed: doomed.length });
    }
    const keptBecause = holdsObject.size;

    if (!previewOnly) {
      toDelete.sort((a, b) => b - a); // from the end, so earlier indexes stay valid
      for (let offset = 0; offset < toDelete.length; offset += DELETE_BATCH) {
        for (const i of toDelete.slice(offset, offset + DELETE_BATCH)) deleteParagraph(items[i]!);
        await context.sync();
      }
    }
    const verb = previewOnly ? "Would remove" : "Removed";
    return {
      summary: toDelete.length === 0
        ? `No runs of blank paragraphs longer than ${max}`
        : `${verb} ${toDelete.length} blank paragraph${toDelete.length === 1 ? "" : "s"} in ${runs.length} run${runs.length === 1 ? "" : "s"} (keeping ${max} per run)`,
      preview: previewOnly,
      ...(previewOnly ? { would_remove: toDelete.length } : { removed: toDelete.length }),
      runs: runs.slice(0, 50),
      runs_truncated: runs.length > 50 ? true : undefined,
      paragraphs_before: items.length,
      kept_because_not_empty: keptBecause,
      ...(!previewOnly && toDelete.length > 0 ? { note: "Paragraph indexes have shifted; read the outline again before using old indexes." } : {}),
    };
  });
}
