import { z } from "zod";

/**
 * Word tools. Paragraph indexes are 0-based positions in `body.paragraphs` (document order, table cell paragraphs
 * included); they are only valid until the document changes, so re-read before relying on an old index.
 */

const Hex = z.string().regex(/^#[0-9A-Fa-f]{6}$/, "must be #RRGGBB");
const ParagraphIndex = z.number().int().min(0).max(1_000_000);
const Occurrence = z.number().int().min(1).max(10_000);

/** Text Word's search can find; longer text is matched through its start and end. */
export const MAX_WORD_TEXT = 100_000;
const Text = z.string().max(MAX_WORD_TEXT);

/* -------------------------------------------------------------------------- */
/* Read                                                                        */
/* -------------------------------------------------------------------------- */

export const GetDocOutlineInputSchema = z.object({
  max_headings: z.number().int().min(1).max(1000).default(300),
}).strict();
export type GetDocOutlineInput = z.infer<typeof GetDocOutlineInputSchema>;

export const ReadDocSectionInputSchema = z.object({
  heading: z.string().min(1).max(500).optional().describe("Read the section under the first heading whose text matches (case-insensitive, trimmed): up to the next heading of the same or a higher level."),
  paragraph_index: ParagraphIndex.optional().describe("Start at this paragraph (0-based). Without heading and paragraph_index, reading starts at the top of the document."),
  count: z.number().int().min(1).max(500).default(40).describe("Most paragraphs to return."),
}).strict();
export type ReadDocSectionInput = z.infer<typeof ReadDocSectionInputSchema>;

export const SearchDocTextInputSchema = z.object({
  query: z.string().min(1).max(255),
  match_case: z.boolean().default(false),
  whole_word: z.boolean().default(false),
  max_results: z.number().int().min(1).max(100).default(20),
}).strict();
export type SearchDocTextInput = z.infer<typeof SearchDocTextInputSchema>;

/* -------------------------------------------------------------------------- */
/* Edit text, lists, blank lines                                               */
/* -------------------------------------------------------------------------- */

export const EditDocTextInputSchema = z.object({
  old_text: Text.min(1).describe("Exact text to replace (a line break is \\n). Use a long enough piece to be unique, or pass occurrence / all."),
  new_text: Text.describe("Replacement text; empty deletes the match."),
  match_case: z.boolean().default(true),
  occurrence: Occurrence.default(1).describe("Which match to replace (1 = first). Ignored with all."),
  all: z.boolean().default(false).describe("Replace every match in the body."),
  track: z.boolean().optional().describe("Turn track changes on for this edit (needs WordApi 1.6). Left as the document has it when omitted."),
}).strict();
export type EditDocTextInput = z.infer<typeof EditDocTextInputSchema>;

export const ListItemSchema = z.union([
  z.string().min(1).max(5000),
  z.object({ text: z.string().min(1).max(5000), level: z.number().int().min(0).max(8).default(0) }).strict(),
]);

export const EditDocListInputSchema = z.object({
  operation: z.enum(["bullets", "numbered"]).describe("Create a new bulleted or numbered list from items."),
  items: z.array(ListItemSchema).min(1).max(200),
  after_paragraph: ParagraphIndex.optional().describe("Insert after this paragraph index; the end of the document when omitted."),
}).strict();
export type EditDocListInput = z.infer<typeof EditDocListInputSchema>;

export const CollapseBlankParagraphsInputSchema = z.object({
  max: z.number().int().min(0).max(5).default(1).describe("Most blank paragraphs allowed in a row; longer runs are cut down to this many."),
  preview: z.boolean().default(false).describe("Only count the paragraphs that would be removed."),
}).strict();
export type CollapseBlankParagraphsInput = z.infer<typeof CollapseBlankParagraphsInputSchema>;

export const ProposedEditSchema = z.object({
  old_text: Text.min(1),
  new_text: Text,
  match_case: z.boolean().default(true),
  occurrence: Occurrence.default(1),
  reason: z.string().max(500).optional(),
  alternatives: z.array(Text).max(4).optional().describe("Other wordings the user may pick instead of new_text."),
}).strict();

export const ProposeDocEditsInputSchema = z.object({
  edits: z.array(ProposedEditSchema).min(1).max(50),
  title: z.string().max(200).optional().describe("Heading of the review card."),
}).strict();
export type ProposeDocEditsInput = z.infer<typeof ProposeDocEditsInputSchema>;

/* -------------------------------------------------------------------------- */
/* Comments and tracked changes                                                */
/* -------------------------------------------------------------------------- */

export const ReadCommentsInputSchema = z.object({
  include_resolved: z.boolean().default(true),
  max: z.number().int().min(1).max(500).default(100),
}).strict();
export type ReadCommentsInput = z.infer<typeof ReadCommentsInputSchema>;

export const AddCommentInputSchema = z.object({
  anchor_text: Text.min(1).describe("Text the comment attaches to."),
  text: z.string().min(1).max(10_000),
  occurrence: Occurrence.default(1),
}).strict();
export type AddCommentInput = z.infer<typeof AddCommentInputSchema>;

export const ReplyCommentInputSchema = z.object({
  comment_id: z.string().min(1).max(200),
  text: z.string().min(1).max(10_000),
}).strict();
export type ReplyCommentInput = z.infer<typeof ReplyCommentInputSchema>;

export const ResolveCommentInputSchema = z.object({
  comment_id: z.string().min(1).max(200),
  resolved: z.boolean().default(true),
}).strict();
export type ResolveCommentInput = z.infer<typeof ResolveCommentInputSchema>;

export const ReadTrackedChangesInputSchema = z.object({
  max: z.number().int().min(1).max(500).default(100),
}).strict();
export type ReadTrackedChangesInput = z.infer<typeof ReadTrackedChangesInputSchema>;

export const AcceptRejectChangesInputSchema = z.object({
  mode: z.enum(["accept", "reject"]),
  scope: z.enum(["all", "selection"]).default("all"),
  author: z.string().min(1).max(200).optional().describe("Only changes by this author."),
}).strict();
export type AcceptRejectChangesInput = z.infer<typeof AcceptRejectChangesInputSchema>;

export const SetTrackChangesInputSchema = z.object({
  enabled: z.boolean(),
  mine_only: z.boolean().default(false).describe("With enabled: track only the current user's edits."),
}).strict();
export type SetTrackChangesInput = z.infer<typeof SetTrackChangesInputSchema>;

/* -------------------------------------------------------------------------- */
/* Formatting, tables, structure                                               */
/* -------------------------------------------------------------------------- */

export const WordFontSchema = z.object({
  name: z.string().max(64).optional(),
  size: z.number().min(1).max(400).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.enum(["None", "Single", "Double"]).optional(),
  color: Hex.optional(),
  highlight: Hex.nullable().optional().describe("Highlight colour; null removes it."),
}).strict();

export const ApplyStyleInputSchema = z.object({
  target: z.enum(["selection", "paragraphs"]).default("selection"),
  paragraph_index: ParagraphIndex.optional().describe("First paragraph for target=paragraphs."),
  count: z.number().int().min(1).max(500).default(1),
  style: z.string().min(1).max(100).optional().describe('Style name ("Heading 1", "Quote") or built-in id (Heading1, Title, ListBullet, Normal).'),
  font: WordFontSchema.optional(),
  alignment: z.enum(["Left", "Centered", "Right", "Justified"]).optional(),
  space_after: z.number().min(0).max(1000).optional(),
  space_before: z.number().min(0).max(1000).optional(),
  line_spacing: z.number().min(0).max(1000).optional(),
}).strict().refine((v) => v.style !== undefined || v.font !== undefined || v.alignment !== undefined || v.space_after !== undefined || v.space_before !== undefined || v.line_spacing !== undefined, { message: "give at least one of style, font, alignment or spacing" })
  .refine((v) => v.target !== "paragraphs" || v.paragraph_index !== undefined, { message: "paragraph_index is required for target paragraphs", path: ["paragraph_index"] });
export type ApplyStyleInput = z.infer<typeof ApplyStyleInputSchema>;

const Cell = z.string().max(5000);

export const InsertTableInputSchema = z.object({
  values: z.array(z.array(Cell).min(1).max(63)).min(1).max(500),
  after_paragraph: ParagraphIndex.optional().describe("Insert after this paragraph; the end of the document when omitted."),
  header_row: z.boolean().default(true),
  style: z.string().max(100).optional().describe('Table style name, e.g. "Grid Table 4 - Accent 1".'),
}).strict().refine((v) => new Set(v.values.map((row) => row.length)).size === 1, { message: "every row needs the same number of cells", path: ["values"] });
export type InsertTableInput = z.infer<typeof InsertTableInputSchema>;

export const EditTableInputSchema = z.object({
  table_index: z.number().int().min(0).max(10_000).describe("0-based position of the table in the document."),
  operation: z.enum(["read", "set_cell", "add_rows", "delete_rows", "add_columns", "delete_columns", "merge", "delete_table", "set_style"]),
  row: z.number().int().min(0).max(100_000).optional(),
  column: z.number().int().min(0).max(1000).optional(),
  count: z.number().int().min(1).max(500).default(1),
  text: z.string().max(5000).optional(),
  to_row: z.number().int().min(0).max(100_000).optional().describe("merge: last row of the area."),
  to_column: z.number().int().min(0).max(1000).optional().describe("merge: last column of the area."),
  position: z.enum(["before", "after"]).default("after"),
  style: z.string().max(100).optional(),
}).strict().superRefine((v, ctx) => {
  const need = (field: keyof typeof v) => { if (v[field] === undefined) ctx.addIssue({ code: "custom", message: `${String(field)} is required for ${v.operation}`, path: [field] }); };
  if (v.operation === "set_cell") { need("row"); need("column"); need("text"); }
  if (v.operation === "add_rows" || v.operation === "delete_rows") need("row");
  if (v.operation === "add_columns" || v.operation === "delete_columns") need("column");
  if (v.operation === "merge") { need("row"); need("column"); need("to_row"); need("to_column"); }
  if (v.operation === "set_style") need("style");
});
export type EditTableInput = z.infer<typeof EditTableInputSchema>;

export const ManageContentControlsInputSchema = z.object({
  operation: z.enum(["list", "set_text", "insert", "delete"]),
  tag: z.string().max(200).optional(),
  title: z.string().max(200).optional(),
  text: z.string().max(20_000).optional(),
}).strict().superRefine((v, ctx) => {
  if ((v.operation === "set_text" || v.operation === "delete") && v.tag === undefined && v.title === undefined) ctx.addIssue({ code: "custom", message: "give tag or title to pick the control", path: ["tag"] });
  if (v.operation === "set_text" && v.text === undefined) ctx.addIssue({ code: "custom", message: "text is required for set_text", path: ["text"] });
  if (v.operation === "insert" && v.tag === undefined && v.title === undefined) ctx.addIssue({ code: "custom", message: "give a tag or a title for the new control", path: ["tag"] });
});
export type ManageContentControlsInput = z.infer<typeof ManageContentControlsInputSchema>;

export const InsertFieldInputSchema = z.object({
  field_type: z.enum(["Page", "NumPages", "Date", "Time", "Author", "Title", "FileName", "Seq", "Ref", "PageRef", "StyleRef", "DocProperty"]),
  text: z.string().max(500).optional().describe("Field instruction text, e.g. the bookmark name for Ref/PageRef or the sequence name for Seq."),
  location: z.enum(["selection", "end"]).default("selection"),
}).strict();
export type InsertFieldInput = z.infer<typeof InsertFieldInputSchema>;

export const UpdateFieldsInputSchema = z.object({}).strict();
export type UpdateFieldsInput = z.infer<typeof UpdateFieldsInputSchema>;

export const InsertTocInputSchema = z.object({
  levels: z.number().int().min(1).max(9).default(3).describe("Heading levels to include (1..levels)."),
  after_paragraph: ParagraphIndex.optional().describe("Insert after this paragraph; at the selection when omitted."),
}).strict();
export type InsertTocInput = z.infer<typeof InsertTocInputSchema>;

export const UpdateTocInputSchema = z.object({}).strict();
export type UpdateTocInput = z.infer<typeof UpdateTocInputSchema>;

export const InsertHyperlinkInputSchema = z.object({
  text: z.string().min(1).max(2000).describe("Link text. With anchor_text, the matching text becomes the link instead."),
  url: z.string().min(1).max(2000).regex(/^(https?:\/\/|mailto:|#)/i, "must start with http://, https://, mailto: or #"),
  anchor_text: Text.optional().describe("Turn this existing text into the link."),
  occurrence: Occurrence.default(1),
}).strict();
export type InsertHyperlinkInput = z.infer<typeof InsertHyperlinkInputSchema>;

export const InsertBookmarkInputSchema = z.object({
  name: z.string().min(1).max(40).regex(/^[A-Za-z][A-Za-z0-9_]*$/, "letters, digits and underscore, starting with a letter"),
  anchor_text: Text.optional().describe("Bookmark this text; the selection when omitted."),
  occurrence: Occurrence.default(1),
}).strict();
export type InsertBookmarkInput = z.infer<typeof InsertBookmarkInputSchema>;

export const ManageHeadersFootersInputSchema = z.object({
  operation: z.enum(["read", "set"]),
  section: z.number().int().min(0).max(1000).default(0),
  part: z.enum(["header", "footer"]).default("header"),
  type: z.enum(["Primary", "FirstPage", "EvenPages"]).default("Primary"),
  text: z.string().max(20_000).optional(),
}).strict().refine((v) => v.operation !== "set" || v.text !== undefined, { message: "text is required for set", path: ["text"] });
export type ManageHeadersFootersInput = z.infer<typeof ManageHeadersFootersInputSchema>;

export const InsertFootnoteInputSchema = z.object({
  anchor_text: Text.min(1).describe("The footnote reference is placed after this text."),
  text: z.string().min(1).max(5000),
  occurrence: Occurrence.default(1),
  kind: z.enum(["footnote", "endnote"]).default("footnote"),
}).strict();
export type InsertFootnoteInput = z.infer<typeof InsertFootnoteInputSchema>;

export const SetPageLayoutInputSchema = z.object({
  section: z.number().int().min(0).max(1000).default(0),
  orientation: z.enum(["Portrait", "Landscape"]).optional(),
  margins: z.object({
    top: z.number().min(0).max(1000).optional(), bottom: z.number().min(0).max(1000).optional(),
    left: z.number().min(0).max(1000).optional(), right: z.number().min(0).max(1000).optional(),
  }).strict().optional().describe("Margins in points."),
  paper: z.enum(["Letter", "A4", "A3", "Legal"]).optional(),
}).strict().refine((v) => v.orientation !== undefined || v.margins !== undefined || v.paper !== undefined, { message: "give orientation, margins or paper" });
export type SetPageLayoutInput = z.infer<typeof SetPageLayoutInputSchema>;

/** Name -> schema, in the order the tools are listed. */
export const WordToolInputSchemas = {
  get_doc_outline: GetDocOutlineInputSchema,
  read_doc_section: ReadDocSectionInputSchema,
  search_doc_text: SearchDocTextInputSchema,
  edit_doc_text: EditDocTextInputSchema,
  edit_doc_list: EditDocListInputSchema,
  collapse_blank_paragraphs: CollapseBlankParagraphsInputSchema,
  propose_doc_edits: ProposeDocEditsInputSchema,
  read_comments: ReadCommentsInputSchema,
  add_comment: AddCommentInputSchema,
  reply_comment: ReplyCommentInputSchema,
  resolve_comment: ResolveCommentInputSchema,
  read_tracked_changes: ReadTrackedChangesInputSchema,
  accept_reject_changes: AcceptRejectChangesInputSchema,
  set_track_changes: SetTrackChangesInputSchema,
  apply_style: ApplyStyleInputSchema,
  insert_table: InsertTableInputSchema,
  edit_table: EditTableInputSchema,
  manage_content_controls: ManageContentControlsInputSchema,
  insert_field: InsertFieldInputSchema,
  update_fields: UpdateFieldsInputSchema,
  insert_toc: InsertTocInputSchema,
  update_toc: UpdateTocInputSchema,
  insert_hyperlink: InsertHyperlinkInputSchema,
  insert_bookmark: InsertBookmarkInputSchema,
  manage_headers_footers: ManageHeadersFootersInputSchema,
  insert_footnote: InsertFootnoteInputSchema,
  set_page_layout: SetPageLayoutInputSchema,
} as const satisfies Record<string, z.ZodTypeAny>;

export type WordToolName = keyof typeof WordToolInputSchemas;
export const WORD_TOOL_NAMES = Object.keys(WordToolInputSchemas) as WordToolName[];

/** Word tools that change the document. */
export const WORD_WRITE_TOOLS: readonly WordToolName[] = [
  "edit_doc_text", "edit_doc_list", "collapse_blank_paragraphs", "propose_doc_edits", "add_comment", "reply_comment", "resolve_comment",
  "accept_reject_changes", "set_track_changes", "apply_style", "insert_table", "edit_table", "manage_content_controls", "insert_field",
  "update_fields", "insert_toc", "update_toc", "insert_hyperlink", "insert_bookmark", "manage_headers_footers", "insert_footnote", "set_page_layout",
];
