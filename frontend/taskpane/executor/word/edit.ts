import type { CollapseBlankParagraphsInput, EditDocListInput, EditDocParagraphsInput, EditDocTextInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { isSupported, WORD_LISTS, WORD_TRACK } from "../../host/requirements.js";
import { assertParagraphIndex, clip, isBlankText, isTracking, preview, requireSupport, TRACKED_DELETE_NOTE, WORD_API_1_3, wordRun } from "./common.js";
import { setStyle } from "./format.js";
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

/** `track: true` needs WordApi 1.6; stop before touching the document when it is missing. */
function requireTracking(track: boolean): void {
  if (track) requireSupport("Editing with Track Changes on", WORD_TRACK, "Edit without track, or ask the person to turn Track Changes on in the Review tab and then edit.");
}

/** Turn Track Changes on for an edit that asked for it. True when it was off before (it is left on, so the person sees it). */
async function startTracking(context: Word.RequestContext, track: boolean): Promise<boolean> {
  if (!track) return false;
  context.document.load("changeTrackingMode");
  await context.sync();
  const wasOff = context.document.changeTrackingMode === "Off";
  context.document.changeTrackingMode = "TrackAll";
  return wasOff;
}

export async function editDocText(input: EditDocTextInput): Promise<unknown> {
  const track = input.track === true;
  requireTracking(track);
  return wordRun(async (context) => {
    const spec: ReplaceSpec = {
      oldText: input.old_text,
      newText: input.new_text,
      matchCase: input.match_case ?? true,
      occurrence: input.occurrence ?? 1,
      all: input.all ?? false,
    };
    const targets = await pickTargets(context, context.document.body, spec.oldText, { matchCase: spec.matchCase, occurrence: spec.occurrence, all: spec.all, what: "old_text" });
    const turnedOn = await startTracking(context, track);
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
/* edit_doc_paragraphs                                                         */
/* -------------------------------------------------------------------------- */

interface NewParagraph { text: string; style?: string }

/** The paragraphs to write, one per line: a \n inside a text starts another paragraph with the same style. */
export function expandParagraphs(items: EditDocParagraphsInput["paragraphs"]): NewParagraph[] {
  return (items ?? []).flatMap((item) => {
    const { text, style } = typeof item === "string" ? { text: item, style: undefined } : item;
    return text.replace(/\r\n?|\u000b/g, "\n").split("\n").map((line) => (style ? { text: line, style } : { text: line }));
  });
}

/** Styles a new paragraph should not inherit when no style was asked for: the next line after a heading is body text. */
const NOT_INHERITED = /^(?:Heading\d|Title|Subtitle|TocHeading)$/;

/** Load what the edit needs to know about these paragraphs (WordApi 1.3), and refuse paragraphs in tables. */
async function inspect(context: Word.RequestContext, items: Word.Paragraph[], firstIndex: number): Promise<Array<{ heading: boolean; list: boolean }>> {
  if (!isSupported(WORD_API_1_3)) return items.map(() => ({ heading: false, list: false }));
  for (const p of items) p.load("styleBuiltIn,isListItem,tableNestingLevel");
  await context.sync();
  const inTable = items.findIndex((p) => p.tableNestingLevel > 0);
  if (inTable !== -1) throw new ToolFailure("invalid_arguments", `paragraph ${firstIndex + inTable} is inside a table; change table cells with edit_table`);
  return items.map((p) => ({ heading: NOT_INHERITED.test(String(p.styleBuiltIn ?? "")), list: p.isListItem }));
}

/** Give a new paragraph its style: the one asked for, or plain Normal when it would otherwise copy a heading or a list item. */
function styleNew(paragraph: Word.Paragraph, wanted: NewParagraph, plain: { heading: boolean; list: boolean }): void {
  if (wanted.style) { setStyle(paragraph, wanted.style); return; }
  if (!isSupported(WORD_API_1_3)) return;
  if (plain.list) paragraph.detachFromList();
  if (plain.heading || plain.list) paragraph.styleBuiltIn = "Normal";
}

function setText(paragraph: Word.Paragraph, text: string): void {
  if (text === "") paragraph.clear();
  else paragraph.insertText(text, "Replace");
}

export async function editDocParagraphs(input: EditDocParagraphsInput): Promise<unknown> {
  const track = input.track === true;
  requireTracking(track);
  const fresh = expandParagraphs(input.paragraphs);
  return wordRun(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const items = paragraphs.items;
    const total = items.length;
    const tracked = track ? { tracked: true } : {};

    if (input.operation === "insert") {
      const at = input.paragraph_index;
      if (at !== undefined) assertParagraphIndex(at, total);
      const before = at !== undefined && input.position === "before";
      // At the end the new text follows the last paragraph, so that is the one whose look it would take.
      const neighbour = items[at ?? total - 1]!;
      const [plain] = await inspect(context, [neighbour], at ?? total - 1);
      const turnedOn = await startTracking(context, track);
      let previous: Word.Paragraph | undefined;
      fresh.forEach((wanted, i) => {
        const created = i === 0
          ? (at === undefined ? body.insertParagraph(wanted.text, "End") : neighbour.insertParagraph(wanted.text, before ? "Before" : "After"))
          : previous!.insertParagraph(wanted.text, "After");
        styleNew(created, wanted, plain!);
        previous = created;
      });
      await context.sync();
      const first = at === undefined ? total : before ? at : at + 1;
      const where = at === undefined ? "at the end" : `${before ? "before" : "after"} paragraph ${at}`;
      return {
        summary: `Inserted ${fresh.length} paragraph${fresh.length === 1 ? "" : "s"} ${where}: “${preview(fresh[0]!.text, 40)}”`,
        first_paragraph_index: first,
        last_paragraph_index: first + fresh.length - 1,
        inserted: fresh.length,
        ...tracked,
        ...(turnedOn ? { track_changes_left_on: true } : {}),
        note: "Indexes of the paragraphs after these have shifted.",
      };
    }

    const start = input.paragraph_index!;
    const count = input.count ?? 1;
    assertParagraphIndex(start, total);
    if (start + count > total) {
      throw new ToolFailure("invalid_arguments", `count ${count} from paragraph ${start} runs past the end of the document, which has ${total} paragraphs (0..${total - 1})`);
    }
    const targets = items.slice(start, start + count);
    await inspect(context, targets, start);
    const turnedOn = await startTracking(context, track);
    const leftOn = turnedOn ? { track_changes_left_on: true } : {};
    const deletesTracked = track || (input.operation === "delete" || fresh.length < count ? await isTracking(context) : false);

    if (input.operation === "delete") {
      for (const paragraph of [...targets].reverse()) deleteParagraph(paragraph);
      await context.sync();
      return {
        summary: `Deleted ${count} paragraph${count === 1 ? "" : "s"} from paragraph ${start}`,
        deleted: count,
        ...(deletesTracked ? { tracked: true } : {}),
        ...leftOn,
        note: deletesTracked ? TRACKED_DELETE_NOTE : "Indexes after these have shifted. The last paragraph of a document cannot be removed, only emptied.",
      };
    }

    // replace: new text into the paragraphs that stay, then grow or shrink the passage.
    const kept = Math.min(count, fresh.length);
    for (let i = 0; i < kept; i++) {
      setText(targets[i]!, fresh[i]!.text);
      if (fresh[i]!.style) setStyle(targets[i]!, fresh[i]!.style!);
    }
    let previous = targets[kept - 1]!;
    for (const wanted of fresh.slice(kept)) {
      previous = previous.insertParagraph(wanted.text, "After");
      if (wanted.style) setStyle(previous, wanted.style);
    }
    for (const extra of targets.slice(kept).reverse()) deleteParagraph(extra);
    await context.sync();
    return {
      summary: `Replaced ${count} paragraph${count === 1 ? "" : "s"} from paragraph ${start} with ${fresh.length}: “${preview(fresh[0]!.text, 40)}”`,
      first_paragraph_index: start,
      last_paragraph_index: start + fresh.length - 1,
      replaced: kept,
      ...(fresh.length > count ? { inserted: fresh.length - count } : {}),
      ...(fresh.length < count ? { deleted: count - fresh.length } : {}),
      ...tracked,
      ...leftOn,
      ...(fresh.length < count && deletesTracked ? { note: TRACKED_DELETE_NOTE } : fresh.length !== count ? { note: "Indexes after these have shifted." } : {}),
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
