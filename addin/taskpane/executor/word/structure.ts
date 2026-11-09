import type {
  InsertBookmarkInput,
  InsertFieldInput,
  InsertFootnoteInput,
  InsertHyperlinkInput,
  InsertTocInput,
  ManageContentControlsInput,
  ManageHeadersFootersInput,
  SetPageLayoutInput,
  UpdateFieldsInput,
  UpdateTocInput,
} from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { isSupported, WORD_BOOKMARK, WORD_FIELDS, WORD_TOC } from "../../host/requirements.js";
import { assertParagraphIndex, clip, fitResult, preview, requireSupport, WORD_API_1_3, WORD_PAGE_SETUP, wordRun } from "./common.js";
import { pickTargets } from "./search.js";

/* -------------------------------------------------------------------------- */
/* Content controls                                                            */
/* -------------------------------------------------------------------------- */

export async function manageContentControls(input: ManageContentControlsInput): Promise<unknown> {
  return wordRun(async (context) => {
    const body = context.document.body;
    const pick = (): Word.ContentControlCollection => (input.tag !== undefined ? body.contentControls.getByTag(input.tag) : body.contentControls.getByTitle(input.title!));
    const props = "items/id,items/tag,items/title,items/text,items/type";

    if (input.operation === "list") {
      const controls = body.contentControls;
      controls.load(props);
      await context.sync();
      const rows = controls.items.map((c) => ({ id: c.id, tag: c.tag, title: c.title, type: c.type, text: clip(c.text, 300).text }));
      return fitResult(rows.length, (n) => ({ total: rows.length, controls: rows.slice(0, n), truncated: n < rows.length }));
    }

    if (input.operation === "insert") {
      const selection = context.document.getSelection();
      const control = selection.insertContentControl();
      if (input.tag !== undefined) control.tag = input.tag;
      if (input.title !== undefined) control.title = input.title;
      if (input.text !== undefined) control.insertText(input.text, "Replace");
      control.load("id,tag,title,type");
      await context.sync();
      return { summary: `Inserted content control “${input.title ?? input.tag}”`, id: control.id, tag: control.tag, title: control.title, type: control.type };
    }

    const found = pick();
    found.load(props);
    await context.sync();
    const label = input.tag !== undefined ? `tag “${input.tag}”` : `title “${input.title}”`;
    if (found.items.length === 0) throw new ToolFailure("range_not_found", `No content control with ${label}. Use operation list to see the controls in the document.`);
    if (input.operation === "set_text") {
      for (const control of found.items) control.insertText(input.text!, "Replace");
      await context.sync();
      return { summary: `Filled ${found.items.length} content control${found.items.length === 1 ? "" : "s"} with ${label}`, updated: found.items.length, text: clip(input.text!, 200).text };
    }
    // delete: the control goes, what is inside it stays.
    for (const control of found.items) control.delete(true);
    await context.sync();
    return { summary: `Removed ${found.items.length} content control${found.items.length === 1 ? "" : "s"} with ${label} (their text was kept)`, deleted: found.items.length };
  });
}

/* -------------------------------------------------------------------------- */
/* Fields                                                                      */
/* -------------------------------------------------------------------------- */

const FIELDS_NEEDING_TEXT: ReadonlySet<string> = new Set(["Seq", "Ref", "PageRef", "StyleRef", "DocProperty"]);

export async function insertField(input: InsertFieldInput): Promise<unknown> {
  requireSupport("Fields", WORD_FIELDS, "Insert it from Word's Insert > Quick Parts > Field instead.");
  if (FIELDS_NEEDING_TEXT.has(input.field_type) && !input.text) {
    throw new ToolFailure("invalid_arguments", `A ${input.field_type} field needs text: ${input.field_type === "Ref" || input.field_type === "PageRef" ? "the bookmark name" : input.field_type === "Seq" ? "the sequence name, e.g. Figure" : input.field_type === "StyleRef" ? "the style name" : "the property name"}.`);
  }
  return wordRun(async (context) => {
    const location = input.location ?? "selection";
    const target = location === "end" ? context.document.body.getRange("End") : context.document.getSelection();
    const field = target.insertField(location === "end" ? "End" : "Start", input.field_type, input.text);
    await context.sync();
    let updateNote: string | undefined;
    try {
      field.updateResult();
      await context.sync();
    } catch {
      updateNote = "The field was inserted but could not be refreshed; use update_fields.";
    }
    const result = field.result;
    field.load("code");
    result.load("text");
    await context.sync();
    return {
      summary: `Inserted a ${input.field_type} field${location === "end" ? " at the end" : ""}`,
      field_type: input.field_type,
      code: clip(field.code, 200).text,
      result: clip(result.text, 200).text,
      ...(updateNote ? { note: updateNote } : {}),
    };
  });
}

export async function updateFields(_input: UpdateFieldsInput): Promise<unknown> {
  requireSupport("Updating fields", WORD_FIELDS, "Select all and press F9 in Word.");
  return wordRun(async (context) => {
    const fields = context.document.body.fields;
    fields.load("items");
    await context.sync();
    const items = fields.items;
    let failed = 0;
    try {
      for (const field of items) field.updateResult();
      await context.sync();
    } catch {
      // One field the host cannot refresh fails the whole batch: retry one by one and count what could not be refreshed.
      failed = 0;
      for (const field of items) {
        try {
          field.updateResult();
          await context.sync();
        } catch {
          failed++;
        }
      }
    }
    return {
      summary: items.length === 0 ? "The document body has no fields" : `Updated ${items.length - failed} of ${items.length} field${items.length === 1 ? "" : "s"}`,
      fields: items.length,
      updated: items.length - failed,
      failed,
      scope: "body",
    };
  });
}

/* -------------------------------------------------------------------------- */
/* Table of contents (WordApiDesktop 1.4: desktop only)                         */
/* -------------------------------------------------------------------------- */

const TOC_HINT = "Word on the web cannot create or update a table of contents through add-ins. Use Word desktop (Windows or Mac), or ask the person to use References > Table of Contents.";

export async function insertToc(input: InsertTocInput): Promise<unknown> {
  requireSupport("Table of contents", WORD_TOC, TOC_HINT);
  return wordRun(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    let range: Word.Range;
    if (input.after_paragraph !== undefined) {
      paragraphs.load("items");
      await context.sync();
      assertParagraphIndex(input.after_paragraph, paragraphs.items.length, "after_paragraph");
      range = paragraphs.items[input.after_paragraph]!.insertParagraph("", "After").getRange("Whole");
    } else {
      range = context.document.getSelection().getRange("Start");
    }
    const levels = input.levels ?? 3;
    context.document.tablesOfContents.add(range, { upperHeadingLevel: 1, lowerHeadingLevel: levels, includePageNumbers: true, rightAlignPageNumbers: true, useHyperlinksOnWeb: true });
    await context.sync();
    const tocs = context.document.tablesOfContents;
    tocs.load("items");
    await context.sync();
    return {
      summary: `Inserted a table of contents (heading levels 1–${levels})`,
      levels,
      tables_of_contents: tocs.items.length,
      note: "It lists paragraphs using Heading styles; run update_toc after the headings change.",
    };
  });
}

export async function updateToc(_input: UpdateTocInput): Promise<unknown> {
  requireSupport("Table of contents", WORD_TOC, TOC_HINT);
  return wordRun(async (context) => {
    const tocs = context.document.tablesOfContents;
    tocs.load("items");
    await context.sync();
    if (tocs.items.length === 0) throw new ToolFailure("range_not_found", "The document has no table of contents. Use insert_toc to add one.");
    // The TOC object can only refresh page numbers; refreshing the entries means updating the TOC field behind it.
    const fieldsOk = isSupported(WORD_FIELDS);
    let full = 0;
    for (const toc of tocs.items) {
      let refreshed = false;
      if (fieldsOk) {
        try {
          const fields = toc.range.fields;
          fields.load("items/type");
          await context.sync();
          const field = fields.items.find((f) => f.type === "TOC") ?? fields.items[0];
          if (field) {
            field.updateResult();
            await context.sync();
            refreshed = true;
          }
        } catch {
          refreshed = false;
        }
      }
      if (refreshed) full++;
      else {
        toc.updatePageNumbers();
        await context.sync();
      }
    }
    const pagesOnly = tocs.items.length - full;
    return {
      summary: `Updated ${tocs.items.length} table${tocs.items.length === 1 ? "" : "s"} of contents`,
      updated: tocs.items.length,
      entries_refreshed: full,
      ...(pagesOnly > 0 ? { note: `${pagesOnly} table${pagesOnly === 1 ? "" : "s"} of contents only had page numbers refreshed; headings added or renamed since are not shown (Word's Update Table > Update entire table does that).` } : {}),
    };
  });
}

/* -------------------------------------------------------------------------- */
/* Hyperlinks, bookmarks, footnotes                                             */
/* -------------------------------------------------------------------------- */

function htmlEscape(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export async function insertHyperlink(input: InsertHyperlinkInput): Promise<unknown> {
  return wordRun(async (context) => {
    const native = isSupported(WORD_API_1_3);
    let linkedText: string;
    if (input.anchor_text !== undefined) {
      const targets = await pickTargets(context, context.document.body, input.anchor_text, { matchCase: true, occurrence: input.occurrence ?? 1, what: "anchor_text" });
      const range = targets.ranges[0]!;
      if (native) range.hyperlink = input.url;
      else {
        range.load("text");
        await context.sync();
        range.insertHtml(`<a href="${htmlEscape(input.url)}">${htmlEscape(range.text)}</a>`, "Replace");
      }
      linkedText = input.anchor_text;
    } else {
      const selection = context.document.getSelection();
      if (native) {
        const inserted = selection.insertText(input.text, "After");
        inserted.hyperlink = input.url;
      } else {
        selection.insertHtml(`<a href="${htmlEscape(input.url)}">${htmlEscape(input.text)}</a>`, "End");
      }
      linkedText = input.text;
    }
    await context.sync();
    return {
      summary: `Linked “${preview(linkedText, 40)}” to ${preview(input.url, 60)}`,
      url: input.url,
      text: clip(linkedText, 200).text,
      method: native ? "hyperlink_property" : "html",
    };
  });
}

export async function insertBookmark(input: InsertBookmarkInput): Promise<unknown> {
  requireSupport("Bookmarks", WORD_BOOKMARK, "Add it in Word with Insert > Bookmark.");
  return wordRun(async (context) => {
    const range = input.anchor_text !== undefined
      ? (await pickTargets(context, context.document.body, input.anchor_text, { matchCase: true, occurrence: input.occurrence ?? 1, what: "anchor_text" })).ranges[0]!
      : context.document.getSelection();
    let existing = false;
    try {
      const names = context.document.body.getRange("Whole").getBookmarks(true, true);
      await context.sync();
      existing = names.value.some((n) => n.toLowerCase() === input.name.toLowerCase());
    } catch {
      // Only used to warn that a bookmark is being moved.
    }
    range.insertBookmark(input.name);
    await context.sync();
    return {
      summary: `${existing ? "Moved" : "Added"} bookmark “${input.name}”${input.anchor_text ? ` on “${preview(input.anchor_text, 40)}”` : " at the selection"}`,
      name: input.name,
      ...(existing ? { replaced_existing: true, note: "A bookmark with this name already existed and was moved here." } : {}),
    };
  });
}

export async function insertFootnote(input: InsertFootnoteInput): Promise<unknown> {
  const kind = input.kind ?? "footnote";
  requireSupport(kind === "endnote" ? "Endnotes" : "Footnotes", WORD_FIELDS, "Insert it in Word with References > Insert Footnote.");
  return wordRun(async (context) => {
    const targets = await pickTargets(context, context.document.body, input.anchor_text, { matchCase: true, occurrence: input.occurrence ?? 1, what: "anchor_text" });
    const range = targets.ranges[0]!;
    if (kind === "endnote") range.insertEndnote(input.text);
    else range.insertFootnote(input.text);
    await context.sync();
    return { summary: `Added a ${kind} after “${preview(input.anchor_text, 40)}”`, kind, anchor_text: clip(input.anchor_text, 200).text, matches: targets.total, text: clip(input.text, 200).text };
  });
}

/* -------------------------------------------------------------------------- */
/* Headers, footers, page layout                                                */
/* -------------------------------------------------------------------------- */

const HEADER_TEXT_LIMIT = 4000;

export async function manageHeadersFooters(input: ManageHeadersFootersInput): Promise<unknown> {
  return wordRun(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    const index = input.section ?? 0;
    if (index >= sections.items.length) {
      throw new ToolFailure("invalid_arguments", `section ${index} does not exist: the document has ${sections.items.length} section${sections.items.length === 1 ? "" : "s"} (0..${sections.items.length - 1}).`);
    }
    const section = sections.items[index]!;
    const part = input.part ?? "header";
    const type = input.type ?? "Primary";
    const body = part === "header" ? section.getHeader(type) : section.getFooter(type);
    const hint = type === "FirstPage" ? "A first-page header/footer only shows when “Different first page” is on for the section." : type === "EvenPages" ? "An even-page header/footer only shows when “Different odd & even pages” is on." : undefined;
    if (input.operation === "read") {
      body.load("text");
      await context.sync();
      const text = clip(body.text, HEADER_TEXT_LIMIT);
      return { section: index, part, type, sections: sections.items.length, text: text.text, truncated: text.truncated, ...(hint ? { note: hint } : {}) };
    }
    body.insertText(input.text!, "Replace");
    await context.sync();
    return {
      summary: `Set the ${type === "Primary" ? "" : `${type} `}${part} of section ${index}`,
      section: index,
      part,
      type,
      text: clip(input.text!, 200).text,
      ...(hint ? { note: hint } : {}),
    };
  });
}

export async function setPageLayout(input: SetPageLayoutInput): Promise<unknown> {
  requireSupport("Page layout", WORD_PAGE_SETUP, "Word on the web cannot change page setup through add-ins; use Word desktop or the Layout tab.");
  return wordRun(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    const index = input.section ?? 0;
    if (index >= sections.items.length) {
      throw new ToolFailure("invalid_arguments", `section ${index} does not exist: the document has ${sections.items.length} section${sections.items.length === 1 ? "" : "s"}.`);
    }
    const setup = sections.items[index]!.pageSetup;
    // Paper first, then orientation (it swaps width and height), then margins.
    if (input.paper !== undefined) setup.paperSize = input.paper;
    if (input.orientation !== undefined) setup.orientation = input.orientation;
    if (input.margins?.top !== undefined) setup.topMargin = input.margins.top;
    if (input.margins?.bottom !== undefined) setup.bottomMargin = input.margins.bottom;
    if (input.margins?.left !== undefined) setup.leftMargin = input.margins.left;
    if (input.margins?.right !== undefined) setup.rightMargin = input.margins.right;
    await context.sync();
    setup.load("orientation,paperSize,topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight");
    await context.sync();
    const applied = [input.paper && `paper ${input.paper}`, input.orientation && input.orientation.toLowerCase(), input.margins && "margins"].filter((x): x is string => Boolean(x));
    return {
      summary: `Page layout of section ${index}: ${applied.join(", ")}`,
      section: index,
      orientation: setup.orientation,
      paper: setup.paperSize,
      margins: { top: setup.topMargin, bottom: setup.bottomMargin, left: setup.leftMargin, right: setup.rightMargin },
      page: { width: setup.pageWidth, height: setup.pageHeight },
      unit: "points",
    };
  });
}
