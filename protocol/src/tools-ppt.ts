import { z } from "zod";

/**
 * PowerPoint tools. A slide is addressed by its Office slide id ("256"); a shape by `ref` = "<slideId>:<shapeId>"
 * ("256:4"), the form list_slide_shapes returns. Positions are 1-based; sizes are in points.
 */

const Hex = z.string().regex(/^#[0-9A-Fa-f]{6}$/, "must be #RRGGBB");
const SlideId = z.string().min(1).max(100).describe('Office slide id, e.g. "256" (see get_presentation_metadata).');
const ShapeRef = z.string().regex(/^[^:\s]+:[^:\s]+$/, 'must look like "<slideId>:<shapeId>", e.g. "256:4"').describe('Shape reference "<slideId>:<shapeId>" from list_slide_shapes.');
const Points = z.number().min(-10_000).max(100_000);
const Size = z.number().min(0).max(100_000);
const Position = z.number().int().min(1).max(10_000);

export const GetPresentationMetadataInputSchema = z.object({}).strict();
export type GetPresentationMetadataInput = z.infer<typeof GetPresentationMetadataInputSchema>;

export const ListLayoutsInputSchema = z.object({}).strict();
export type ListLayoutsInput = z.infer<typeof ListLayoutsInputSchema>;

export const AddSlideInputSchema = z.object({
  layout_id: z.string().min(1).max(100),
  master_id: z.string().min(1).max(100).optional().describe("Slide master the layout belongs to; found from layout_id when omitted."),
  position: Position.optional().describe("1-based position of the new slide; the end when omitted."),
}).strict();
export type AddSlideInput = z.infer<typeof AddSlideInputSchema>;

export const ListSlideShapesInputSchema = z.object({
  slide_id: SlideId.optional().describe("The selected slide when omitted."),
  include_text: z.boolean().default(true),
}).strict();
export type ListSlideShapesInput = z.infer<typeof ListSlideShapesInputSchema>;

export const ReadSlideTextInputSchema = z.object({
  ref: ShapeRef.optional(),
  slide_id: SlideId.optional().describe("Read every shape of this slide when ref is omitted."),
}).strict().refine((v) => v.ref !== undefined || v.slide_id !== undefined, { message: "give ref or slide_id" });
export type ReadSlideTextInput = z.infer<typeof ReadSlideTextInputSchema>;

export const SlideTextEditSchema = z.object({ ref: ShapeRef, text: z.string().max(50_000) }).strict();

export const EditSlideTextInputSchema = z.object({
  ref: ShapeRef.optional(),
  text: z.string().max(50_000).optional().describe("New text of the shape; line breaks are \\n."),
  edits: z.array(SlideTextEditSchema).min(1).max(50).optional().describe("Several shapes at once."),
}).strict().refine((v) => (v.ref !== undefined && v.text !== undefined) !== (v.edits !== undefined), { message: "give either ref and text, or edits" });
export type EditSlideTextInput = z.infer<typeof EditSlideTextInputSchema>;

export const EditSlideXmlInputSchema = z.object({
  slide_id: SlideId,
  code: z.string().min(1).max(30_000).describe(
    "Body of an async function with `xml` (the slide's parsed XML Document, ppt/slides/slide1.xml), `files` (names of every part in the exported one-slide file), " +
    "`ns` ({ a, p, r } namespace URIs) and `helpers` ({ el(name, attrs?, ...children), q(parent, selector) }). Change `xml` in place and `return` a short summary. " +
    "The slide is exported, changed and inserted back at the same position; the old slide is removed after that.",
  ),
  description: z.string().max(200).optional(),
}).strict();
export type EditSlideXmlInput = z.infer<typeof EditSlideXmlInputSchema>;

const Geometry = { left: Points, top: Points, width: Size, height: Size };
const TextStyle = z.object({
  font: z.string().max(64).optional(), size: z.number().min(1).max(400).optional(), bold: z.boolean().optional(),
  italic: z.boolean().optional(), color: Hex.optional(), align: z.enum(["Left", "Center", "Right", "Justify"]).optional(),
}).strict();

export const SlideElementSchema = z.discriminatedUnion("type", [
  z.object({ type: z.literal("textbox"), text: z.string().max(50_000), ...Geometry, style: TextStyle.optional(), name: z.string().max(100).optional() }).strict(),
  z.object({
    type: z.literal("shape"), shape: z.string().regex(/^[A-Za-z]{2,40}$/).default("Rectangle").describe("PowerPoint.GeometricShapeType name, e.g. Rectangle, RoundRectangle, Ellipse, RightArrow."),
    text: z.string().max(50_000).optional(), ...Geometry, fill: Hex.optional(), line: Hex.optional(), line_weight: z.number().min(0).max(100).optional(),
    style: TextStyle.optional(), name: z.string().max(100).optional(),
  }).strict(),
  z.object({
    type: z.literal("line"), left: Points, top: Points, width: Size, height: Size, line: Hex.optional(), line_weight: z.number().min(0).max(100).optional(),
    connector: z.enum(["Straight", "Elbow", "Curve"]).default("Straight"), name: z.string().max(100).optional(),
  }).strict(),
  z.object({
    type: z.literal("table"), values: z.array(z.array(z.string().max(5000)).min(1).max(30)).min(1).max(100),
    left: Points.optional(), top: Points.optional(), width: Size.optional(), height: Size.optional(), name: z.string().max(100).optional(),
  }).strict(),
]);

export const InsertSlideElementInputSchema = z.object({ slide_id: SlideId, element: SlideElementSchema }).strict()
  .refine((v) => v.element.type !== "table" || new Set(v.element.values.map((row) => row.length)).size === 1, { message: "every row needs the same number of cells", path: ["element", "values"] });
export type InsertSlideElementInput = z.infer<typeof InsertSlideElementInputSchema>;

export const RemoveSlideElementInputSchema = z.object({
  slide_id: SlideId,
  shape_ids: z.array(z.string().min(1).max(100)).min(1).max(100),
}).strict();
export type RemoveSlideElementInput = z.infer<typeof RemoveSlideElementInputSchema>;

export const DuplicateSlideInputSchema = z.object({
  slide_id: SlideId,
  position: Position.optional().describe("1-based position of the copy; right after the original when omitted."),
}).strict();
export type DuplicateSlideInput = z.infer<typeof DuplicateSlideInputSchema>;

export const MoveSlideInputSchema = z.object({ slide_id: SlideId, position: Position }).strict();
export type MoveSlideInput = z.infer<typeof MoveSlideInputSchema>;

export const ModifyPresentationStructureInputSchema = z.object({
  operation: z.enum(["add", "delete", "move"]),
  slide_id: SlideId.optional(),
  slide_ids: z.array(SlideId).min(1).max(100).optional().describe("delete: several slides at once."),
  layout_id: z.string().min(1).max(100).optional(),
  position: Position.optional(),
  confirm_delete: z.boolean().default(false).describe("Must be true to delete slides."),
}).strict().superRefine((v, ctx) => {
  if (v.operation === "add" && v.layout_id === undefined) ctx.addIssue({ code: "custom", message: "layout_id is required for add", path: ["layout_id"] });
  if (v.operation === "move" && (v.slide_id === undefined || v.position === undefined)) ctx.addIssue({ code: "custom", message: "slide_id and position are required for move", path: ["slide_id"] });
  if (v.operation === "delete" && v.slide_id === undefined && v.slide_ids === undefined) ctx.addIssue({ code: "custom", message: "give slide_id or slide_ids for delete", path: ["slide_id"] });
});
export type ModifyPresentationStructureInput = z.infer<typeof ModifyPresentationStructureInputSchema>;

export const VerifySlidesInputSchema = z.object({ slide_ids: z.array(SlideId).min(1).max(200).optional() }).strict();
export type VerifySlidesInput = z.infer<typeof VerifySlidesInputSchema>;

export const VerifySlideVisualInputSchema = z.object({ slide_id: SlideId, width: z.number().int().min(200).max(2000).default(960) }).strict();
export type VerifySlideVisualInput = z.infer<typeof VerifySlideVisualInputSchema>;

export const ExportSlideImageInputSchema = z.object({ slide_id: SlideId, width: z.number().int().min(200).max(2000).default(960) }).strict();
export type ExportSlideImageInput = z.infer<typeof ExportSlideImageInputSchema>;

export const InsertSlideTableInputSchema = z.object({
  slide_id: SlideId,
  values: z.array(z.array(z.string().max(5000)).min(1).max(30)).min(1).max(100),
  left: Points.optional(), top: Points.optional(), width: Size.optional(), height: Size.optional(),
  name: z.string().max(100).optional(),
}).strict().refine((v) => new Set(v.values.map((row) => row.length)).size === 1, { message: "every row needs the same number of cells", path: ["values"] });
export type InsertSlideTableInput = z.infer<typeof InsertSlideTableInputSchema>;

export const EditSlideTableInputSchema = z.object({
  ref: ShapeRef,
  operation: z.enum(["read", "set_cell", "add_rows", "delete_rows", "add_columns", "delete_columns"]),
  row: z.number().int().min(0).max(100_000).optional(),
  column: z.number().int().min(0).max(1000).optional(),
  count: z.number().int().min(1).max(100).default(1),
  text: z.string().max(5000).optional(),
}).strict().superRefine((v, ctx) => {
  const need = (field: "row" | "column" | "text") => { if (v[field] === undefined) ctx.addIssue({ code: "custom", message: `${field} is required for ${v.operation}`, path: [field] }); };
  if (v.operation === "set_cell") { need("row"); need("column"); need("text"); }
  if (v.operation === "add_rows" || v.operation === "delete_rows") need("row");
  if (v.operation === "add_columns" || v.operation === "delete_columns") need("column");
});
export type EditSlideTableInput = z.infer<typeof EditSlideTableInputSchema>;

export const SetSlideBackgroundInputSchema = z.object({ slide_id: SlideId, color: Hex }).strict();
export type SetSlideBackgroundInput = z.infer<typeof SetSlideBackgroundInputSchema>;

export const SetShapeAltTextInputSchema = z.object({
  ref: ShapeRef, title: z.string().max(200).optional(), description: z.string().max(2000).optional(),
}).strict().refine((v) => v.title !== undefined || v.description !== undefined, { message: "give title or description" });
export type SetShapeAltTextInput = z.infer<typeof SetShapeAltTextInputSchema>;

export const SetSlideHyperlinkInputSchema = z.object({
  ref: ShapeRef,
  url: z.string().min(1).max(2000).regex(/^(https?:\/\/|mailto:)/i, "must start with http://, https:// or mailto:"),
  tooltip: z.string().max(200).optional(),
}).strict();
export type SetSlideHyperlinkInput = z.infer<typeof SetSlideHyperlinkInputSchema>;

export const SetShapeZOrderInputSchema = z.object({
  ref: ShapeRef, action: z.enum(["bring_to_front", "send_to_back", "bring_forward", "send_backward"]),
}).strict();
export type SetShapeZOrderInput = z.infer<typeof SetShapeZOrderInputSchema>;

export const GroupShapesInputSchema = z.object({
  slide_id: SlideId, shape_ids: z.array(z.string().min(1).max(100)).min(2).max(200), name: z.string().max(100).optional(),
}).strict();
export type GroupShapesInput = z.infer<typeof GroupShapesInputSchema>;

export const UngroupShapesInputSchema = z.object({ ref: ShapeRef }).strict();
export type UngroupShapesInput = z.infer<typeof UngroupShapesInputSchema>;

export const AlignShapesInputSchema = z.object({
  slide_id: SlideId, shape_ids: z.array(z.string().min(1).max(100)).min(1).max(200),
  align: z.enum(["left", "center", "right", "top", "middle", "bottom"]).optional(),
  distribute: z.enum(["horizontal", "vertical"]).optional(),
  relative_to: z.enum(["shapes", "slide"]).default("shapes"),
}).strict().refine((v) => v.align !== undefined || v.distribute !== undefined, { message: "give align or distribute" });
export type AlignShapesInput = z.infer<typeof AlignShapesInputSchema>;

/** Name -> schema, in the order the tools are listed. */
export const PptToolInputSchemas = {
  get_presentation_metadata: GetPresentationMetadataInputSchema,
  list_layouts: ListLayoutsInputSchema,
  add_slide: AddSlideInputSchema,
  list_slide_shapes: ListSlideShapesInputSchema,
  read_slide_text: ReadSlideTextInputSchema,
  edit_slide_text: EditSlideTextInputSchema,
  edit_slide_xml: EditSlideXmlInputSchema,
  insert_slide_element: InsertSlideElementInputSchema,
  remove_slide_element: RemoveSlideElementInputSchema,
  duplicate_slide: DuplicateSlideInputSchema,
  move_slide: MoveSlideInputSchema,
  modify_presentation_structure: ModifyPresentationStructureInputSchema,
  verify_slides: VerifySlidesInputSchema,
  verify_slide_visual: VerifySlideVisualInputSchema,
  export_slide_image: ExportSlideImageInputSchema,
  insert_slide_table: InsertSlideTableInputSchema,
  edit_slide_table: EditSlideTableInputSchema,
  set_slide_background: SetSlideBackgroundInputSchema,
  set_shape_alt_text: SetShapeAltTextInputSchema,
  set_slide_hyperlink: SetSlideHyperlinkInputSchema,
  set_shape_zorder: SetShapeZOrderInputSchema,
  group_shapes: GroupShapesInputSchema,
  ungroup_shapes: UngroupShapesInputSchema,
  align_shapes: AlignShapesInputSchema,
} as const satisfies Record<string, z.ZodTypeAny>;

export type PptToolName = keyof typeof PptToolInputSchemas;
export const PPT_TOOL_NAMES = Object.keys(PptToolInputSchemas) as PptToolName[];

/** PowerPoint tools that change the presentation. */
export const PPT_WRITE_TOOLS: readonly PptToolName[] = [
  "add_slide", "edit_slide_text", "edit_slide_xml", "insert_slide_element", "remove_slide_element", "duplicate_slide", "move_slide",
  "modify_presentation_structure", "insert_slide_table", "edit_slide_table", "set_slide_background", "set_shape_alt_text", "set_slide_hyperlink",
  "set_shape_zorder", "group_shapes", "ungroup_shapes", "align_shapes",
];
