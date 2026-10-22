import { z } from "zod";
import { CellValueSchema } from "./tools.js";

export const SheetInfoSchema = z.object({
  id: z.string(),
  name: z.string(),
  visibility: z.enum(["Visible", "Hidden", "VeryHidden"]),
  usedRange: z.string().nullable(),
  isActive: z.boolean(),
});
export type SheetInfo = z.infer<typeof SheetInfoSchema>;

export const GetSheetsResultSchema = z.object({
  sheets: z.array(SheetInfoSchema),
  activeSheet: z.string(),
});
export type GetSheetsResult = z.infer<typeof GetSheetsResultSchema>;

export const ShapeInfoSchema = z.object({
  id: z.string().optional(),
  parentGroupId: z.string().optional(),
  rotation: z.number().optional(),
  name: z.string(),
  type: z.string(),
  text: z.string().nullable(),
  textReadStatus: z.enum(["read", "empty", "not_applicable", "unavailable"]).optional(),
  readWarnings: z.array(z.object({ property: z.string(), code: z.string(), message: z.string() })).optional(),
  left: z.number(),
  top: z.number(),
  width: z.number(),
  height: z.number(),
  connection: z.object({
    from: z.object({ id: z.string(), name: z.string(), site: z.number() }).nullable(),
    to: z.object({ id: z.string(), name: z.string(), site: z.number() }).nullable(),
    connectorType: z.string(),
    beginArrow: z.string(),
    endArrow: z.string(),
  }).optional(),
});
export type ShapeInfo = z.infer<typeof ShapeInfoSchema>;

export const SelectionSnapshotSchema = z.object({
  workbookName: z.string(),
  workbookUrl: z.string().optional(),
  sheet: z.string(),
  sheetId: z.string().nullable(),
  address: z.string(),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  /** Values of the selection, truncated to the configured limits. */
  values: z.array(z.array(CellValueSchema)),
  /** Column headers of the surrounding table/used range, when detected. */
  headers: z.array(z.string()).nullable(),
  truncated: z.boolean(),
  /** Set when the selection is not a cell range (a shape or chart is selected). */
  note: z.string().optional(),
  activeShape: ShapeInfoSchema.optional(),
  /** COM IDs are native IDs; use names when targeting Office.js shape tools. */
  selectedShapes: z.array(z.object({ name: z.string(), nativeId: z.string(), nativeType: z.number().int(), connector: z.boolean() })).optional(),
  selectedShapesSource: z.literal("windows_com").optional(),
  selectedShapesStatus: z.enum(["complete", "unavailable"]).optional(),
  selectedShapesNote: z.string().optional(),
  shapeSelectionStatus: z.enum(["active", "unavailable", "none"]).optional(),
});
export type SelectionSnapshot = z.infer<typeof SelectionSnapshotSchema>;

export const GetSelectionResultSchema = SelectionSnapshotSchema;
export type GetSelectionResult = z.infer<typeof GetSelectionResultSchema>;

export const ReadRangeResultSchema = z.object({
  sheet: z.string(),
  address: z.string(),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  values: z.array(z.array(CellValueSchema)),
  formulas: z.array(z.array(z.string())).nullable().optional(),
  /** Present instead of `values` when `format: "csv"` was requested. */
  csv: z.string().optional(),
  truncated: z.boolean(),
});
export type ReadRangeResult = z.infer<typeof ReadRangeResultSchema>;

export const WriteRangeResultSchema = z.object({
  sheet: z.string(),
  address: z.string(),
  rowsWritten: z.number().int().nonnegative(),
  columnsWritten: z.number().int().nonnegative(),
  notesWritten: z.number().int().nonnegative().optional(),
});
export type WriteRangeResult = z.infer<typeof WriteRangeResultSchema>;

export const SetFormulaResultSchema = z.object({
  sheet: z.string(),
  address: z.string(),
  formula: z.string(),
});
export type SetFormulaResult = z.infer<typeof SetFormulaResultSchema>;

export const FormatRangeResultSchema = z.object({
  sheet: z.string(),
  address: z.string(),
  applied: z.array(z.string()),
});
export type FormatRangeResult = z.infer<typeof FormatRangeResultSchema>;

export const ClearRangeResultSchema = z.object({
  sheet: z.string(),
  address: z.string(),
  clearType: z.enum(["contents", "formats", "all"]),
});

export const SearchDataResultSchema = z.object({
  matches: z.array(z.object({ sheet: z.string(), address: z.string(), value: CellValueSchema })),
  totalFound: z.number().int().nonnegative(),
  returned: z.number().int().nonnegative(),
  hasMore: z.boolean(),
  nextOffset: z.number().int().nullable(),
  /** True when a size limit stopped the scan before every cell was checked. */
  scanTruncated: z.boolean(),
});

export const ModifyStructureResultSchema = z.object({
  sheet: z.string(),
  operation: z.string(),
  target: z.string(),
});

export const InsertRowsResultSchema = z.object({
  sheet: z.string(),
  at: z.string(),
  count: z.number().int().positive(),
  shiftedRange: z.string(),
});
export type InsertRowsResult = z.infer<typeof InsertRowsResultSchema>;

export const ToolErrorSchema = z.object({
  code: z.enum([
    "invalid_arguments",
    "unknown_tool",
    "sheet_not_found",
    "range_not_found",
    "would_overwrite",
    "confirmation_required",
    "unsupported",
    "workbook_readonly",
    "size_limit_exceeded",
    "office_error",
    "timeout",
    "cancelled",
    "disconnected",
    "internal_error",
  ]),
  message: z.string(),
  /** Optional structured detail safe to surface to the model. */
  detail: z.unknown().optional(),
});
export type ToolError = z.infer<typeof ToolErrorSchema>;

export const ToolResultSchema = z.discriminatedUnion("ok", [
  z.object({ ok: z.literal(true), data: z.unknown() }),
  z.object({ ok: z.literal(false), error: ToolErrorSchema }),
]);
export type ToolResult = z.infer<typeof ToolResultSchema>;

/** Office.js executor returns this shape for every tool call. */
export const ExecutedToolResultSchema = z.object({
  callId: z.string().min(1),
  result: ToolResultSchema,
});
export type ExecutedToolResult = z.infer<typeof ExecutedToolResultSchema>;
