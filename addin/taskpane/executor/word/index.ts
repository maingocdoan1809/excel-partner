import { WordToolInputSchemas, type ToolResult, type WordToolName } from "../../../../protocol/src/index.js";
import { ToolFailure, toToolError } from "../common.js";
import { addComment, readComments, replyComment, resolveComment } from "./comments.js";
import { assertWithinLimit, asWordFailure } from "./common.js";
import { collapseBlankParagraphs, editDocList, editDocText } from "./edit.js";
import { applyStyle } from "./format.js";
import { proposeDocEdits } from "./proposal.js";
import { getDocOutline, readDocSection, searchDocText } from "./read.js";
import { getWordSelection } from "./selection.js";
import {
  insertBookmark, insertField, insertFootnote, insertHyperlink, insertToc, manageContentControls, manageHeadersFooters,
  setPageLayout, updateFields, updateToc,
} from "./structure.js";
import { editTable, insertTable } from "./tables.js";
import { acceptRejectChanges, readTrackedChanges, setTrackChanges } from "./track.js";

export { setProposalReviewer, hasProposalReviewer } from "./proposal.js";
export type { ProposalItem, ProposalRequest, ProposalResponse, ProposalReviewer } from "./proposal.js";
export { getWordSelection, getWordSelectionLabel, readWordDocumentName } from "./selection.js";

class UnknownTool extends Error {}

/** Word tools that change the document; the pane lists their results as changes. */
export { WORD_WRITE_TOOLS } from "../../../../protocol/src/index.js";

function isWordTool(tool: string): tool is WordToolName {
  return Object.prototype.hasOwnProperty.call(WordToolInputSchemas, tool);
}

type Handler = (args: any, signal?: AbortSignal) => Promise<unknown>;

const HANDLERS: Record<WordToolName, Handler> = {
  get_doc_outline: getDocOutline,
  read_doc_section: readDocSection,
  search_doc_text: searchDocText,
  edit_doc_text: editDocText,
  edit_doc_list: editDocList,
  collapse_blank_paragraphs: collapseBlankParagraphs,
  propose_doc_edits: proposeDocEdits,
  read_comments: readComments,
  add_comment: addComment,
  reply_comment: replyComment,
  resolve_comment: resolveComment,
  read_tracked_changes: readTrackedChanges,
  accept_reject_changes: acceptRejectChanges,
  set_track_changes: setTrackChanges,
  apply_style: applyStyle,
  insert_table: insertTable,
  edit_table: editTable,
  manage_content_controls: manageContentControls,
  insert_field: insertField,
  update_fields: updateFields,
  insert_toc: insertToc,
  update_toc: updateToc,
  insert_hyperlink: insertHyperlink,
  insert_bookmark: insertBookmark,
  manage_headers_footers: manageHeadersFooters,
  insert_footnote: insertFootnote,
  set_page_layout: setPageLayout,
};

async function run(tool: string, args: unknown, signal?: AbortSignal): Promise<unknown> {
  if (tool === "get_selection") return getWordSelection();
  if (!isWordTool(tool)) throw new UnknownTool(`'${tool}' is not a Word tool`);
  // The bridge validated the arguments already; parsing again fills defaults and keeps a bad call from reaching Word.
  const parsed = WordToolInputSchemas[tool].safeParse(args ?? {});
  if (!parsed.success) {
    const message = parsed.error.issues.map((i) => `${i.path.join(".") || "arguments"}: ${i.message}`).join("; ");
    throw new ToolFailure("invalid_arguments", message);
  }
  return HANDLERS[tool](parsed.data, signal);
}

/**
 * Execute one Word tool request inside the open document through Word.run. Same contract as the Excel `executeTool`:
 * success is `{ ok: true, data }`, anything the executor or Word raises becomes a `ToolError`, an unknown name is
 * `unknown_tool`. Every write tool's data has a short `summary` for the Changes list.
 */
export async function executeWordTool(tool: string, args: unknown, signal?: AbortSignal): Promise<ToolResult> {
  try {
    const data = await run(tool, args, signal);
    assertWithinLimit(data);
    return { ok: true, data };
  } catch (error) {
    if (error instanceof UnknownTool) return { ok: false, error: { code: "unknown_tool", message: error.message } };
    return { ok: false, error: toToolError(asWordFailure(error)) };
  }
}
