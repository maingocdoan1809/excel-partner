import { z } from "zod";

/** A single spreadsheet cell value as seen by Office.js. */
export const CellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);
export type CellValue = z.infer<typeof CellValueSchema>;

/** A1-style address, e.g. `A1`, `B2:D10`, or a table column range. */
export const A1AddressSchema = z
  .string()
  .min(1)
  .max(256)
  .regex(/^[A-Za-z$][A-Za-z0-9$]*(:[A-Za-z$][A-Za-z0-9$]*)?$/, "must be an A1 address");

/** A1 address, whole-column (`D:D`) or whole-row (`3:3`) reference. */
export const RangeRefSchema = z
  .string()
  .min(1)
  .max(256)
  .regex(
    /^(\$?[A-Za-z]{1,3}\$?\d{0,7}(:\$?[A-Za-z]{1,3}\$?\d{0,7})?|\$?\d{1,7}:\$?\d{1,7})$/,
    "must be an A1 range, a column range like D:D or a row range like 3:3",
  );

export const SheetRefSchema = z.object({
  /** Sheet display name. Mutually exclusive with `sheetId` in practice. */
  name: z.string().min(1).max(255).optional(),
  /** Sheet id from Office.js; survives renames. */
  id: z.string().min(1).max(255).optional(),
});
export type SheetRef = z.infer<typeof SheetRefSchema>;

const AlignmentSchema = z.enum(["Left", "Center", "Right", "Justify", "General"]);
const VAlignmentSchema = z.enum(["Top", "Center", "Bottom"]);

export const RangeFormatSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontSize: z.number().min(1).max(409).optional(),
  fontColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
  fillColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
  numberFormat: z.string().max(255).optional(),
  horizontalAlignment: AlignmentSchema.optional(),
  verticalAlignment: VAlignmentSchema.optional(),
  wrapText: z.boolean().optional(),
  columnWidth: z.number().positive().max(255).optional(),
  rowHeight: z.number().positive().max(409).optional(),
  /** `all` = every cell edge, `outline` = outer edge only, `none` = remove. */
  borders: z.enum(["all", "outline", "none"]).optional(),
  /** Fit column widths to their content. */
  autofitColumns: z.boolean().optional(),
});
export type RangeFormat = z.infer<typeof RangeFormatSchema>;

export const GetSelectionInputSchema = z.object({});
export const GetSheetsInputSchema = z.object({});

export const ReadRangeInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  address: A1AddressSchema,
  /** Also return the formulas of the range (parallel to `values`). */
  includeFormulas: z.boolean().default(false),
  /** `grid` returns a 2D array; `csv` returns compact CSV text (cheaper for large reads). */
  format: z.enum(["grid", "csv"]).default("grid"),
  /** Return at most this many rows (paging with `offset`); default is the executor cap. */
  maxRows: z.number().int().min(1).max(20_000).optional(),
  /** Rows to skip from the top of the range. */
  offset: z.number().int().min(0).default(0),
});
export type ReadRangeInput = z.infer<typeof ReadRangeInputSchema>;

export const WriteRangeInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  address: A1AddressSchema,
  values: z.array(z.array(CellValueSchema)).min(1),
  /**
   * `values`: the address must be one cell or exactly match the size of `values`.
   * `top_left`: `address` is only the anchor; the block extends to fit `values`.
   */
  mode: z.enum(["values", "top_left"]).default("values"),
  /**
   * Writing over non-empty cells is refused unless this is true. Set it only
   * after the user confirmed, or when they explicitly asked to replace data.
   */
  allowOverwrite: z.boolean().default(false),
  /** Cell notes to attach, keyed by cell address (e.g. `{ "B2": "Source: ..." }`). */
  notes: z.record(z.string().max(32), z.string().max(2000)).optional(),
});
export type WriteRangeInput = z.infer<typeof WriteRangeInputSchema>;

export const SetFormulaInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  /**
   * One cell, or a range. The formula is written to the top-left cell and then
   * filled across the range, so relative references adjust per cell.
   */
  address: A1AddressSchema,
  formula: z.string().min(1).max(8192),
  /** Formula notation language; Office.js defaults to the UI language. */
  language: z.literal("A1").default("A1"),
  allowOverwrite: z.boolean().default(false),
});
export type SetFormulaInput = z.infer<typeof SetFormulaInputSchema>;

export const FormatRangeInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  address: A1AddressSchema,
  format: RangeFormatSchema,
});
export type FormatRangeInput = z.infer<typeof FormatRangeInputSchema>;

export const ClearRangeInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  address: RangeRefSchema,
  /** `contents` keeps formatting; `formats` keeps values; `all` removes both. */
  clearType: z.enum(["contents", "formats", "all"]).default("contents"),
});
export type ClearRangeInput = z.infer<typeof ClearRangeInputSchema>;

export const SearchDataInputSchema = z.object({
  /** Limit to one sheet; omitted searches every visible sheet. */
  sheet: SheetRefSchema.optional(),
  /** Limit to an A1 range on that sheet. */
  address: A1AddressSchema.optional(),
  searchTerm: z.string().min(1).max(500),
  matchCase: z.boolean().default(false),
  /** Match only cells whose whole value equals the term. */
  wholeCell: z.boolean().default(false),
  limit: z.number().int().min(1).max(500).default(50),
  offset: z.number().int().min(0).default(0),
});
export type SearchDataInput = z.infer<typeof SearchDataInputSchema>;

export const ModifyStructureInputSchema = z
  .object({
    sheet: SheetRefSchema.optional(),
    operation: z.enum(["insert", "delete", "hide", "unhide", "freeze", "unfreeze", "group", "ungroup", "show_levels"]),
    dimension: z.enum(["rows", "columns"]).optional(),
    /** Row number (`5`) or column letters (`C`). Not used by freeze/unfreeze. */
    reference: z.string().regex(/^([A-Za-z]{1,3}|\d{1,7})$/).optional(),
    /** Rows/columns affected; for `freeze` the leading count; for `show_levels` the outline level to show (1 = top level only). */
    count: z.number().int().min(1).max(1000).default(1),
    position: z.enum(["before", "after"]).default("before"),
  })
  .superRefine((value, ctx) => {
    const needsReference = ["insert", "delete", "hide", "unhide", "group", "ungroup"].includes(value.operation);
    if (value.operation !== "unfreeze" && !value.dimension) {
      ctx.addIssue({ code: "custom", path: ["dimension"], message: "dimension is required" });
    }
    if (needsReference && !value.reference) {
      ctx.addIssue({ code: "custom", path: ["reference"], message: "reference is required" });
    }
    if (!needsReference && value.reference) {
      ctx.addIssue({ code: "custom", path: ["reference"], message: "reference is not used by freeze/unfreeze/show_levels" });
    }
    if (value.operation === "show_levels" && value.count > 8) {
      ctx.addIssue({ code: "custom", path: ["count"], message: "outline level must be 1-8" });
    }
    if (value.reference && value.dimension) {
      const isRow = /^\d+$/.test(value.reference);
      if ((value.dimension === "rows") !== isRow) {
        ctx.addIssue({
          code: "custom",
          path: ["reference"],
          message: value.dimension === "rows" ? "rows need a row number" : "columns need column letters",
        });
      }
    }
  });
export type ModifyStructureInput = z.infer<typeof ModifyStructureInputSchema>;

/**
 * Shapes whose connection-site layout is known: four sites (top, left, bottom,
 * right) or eight (the ellipse family, clockwise from the top-left going down).
 */
export const FOUR_SITE_SHAPES = [
  "Rectangle",
  "RoundRectangle",
  "Diamond",
  "FlowChartProcess",
  "FlowChartAlternateProcess",
  "FlowChartDecision",
  "FlowChartInputOutput",
  "FlowChartPredefinedProcess",
  "FlowChartDocument",
  "FlowChartTerminator",
  "FlowChartPreparation",
  "FlowChartManualInput",
  "FlowChartManualOperation",
  "FlowChartDelay",
  "FlowChartDisplay",
  "FlowChartMagneticDisk",
] as const;
export const EIGHT_SITE_SHAPES = ["Ellipse", "FlowChartConnector", "FlowChartOr", "FlowChartSummingJunction"] as const;

export const ShapeTypeSchema = z.enum([...FOUR_SITE_SHAPES, ...EIGHT_SITE_SHAPES, "TextBox"]);
export type ShapeType = z.infer<typeof ShapeTypeSchema>;

const ShapeStyleSchema = z.object({
  fillColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
  /** Outline colour; use `none` for no outline. */
  lineColor: z.union([z.string().regex(/^#[0-9A-Fa-f]{6}$/), z.literal("none")]).optional(),
  /** Outline weight in points; Excel only takes whole points, so fractions are rounded. */
  lineWeight: z.number().min(0.25).max(20).optional(),
  fontColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
  fontSize: z.number().min(1).max(409).optional(),
  fontName: z.string().max(64).optional(),
  bold: z.boolean().optional(),
  horizontalAlignment: z.enum(["Left", "Center", "Right"]).optional(),
});

const ShapeNameSchema = z.string().min(1).max(64);

/** Where a shape goes: over a cell range (`anchor`), or in absolute points. */
const ShapePlacementSchema = z.object({
  anchor: RangeRefSchema.optional(),
  left: z.number().min(0).max(1_000_000).optional(),
  top: z.number().min(0).max(1_000_000).optional(),
  width: z.number().min(1).max(100_000).optional(),
  height: z.number().min(1).max(100_000).optional(),
});

export const AddShapeInputSchema = z
  .object({
    sheet: SheetRefSchema.optional(),
    /** Unique name on the sheet; connectors refer to shapes by name. */
    name: ShapeNameSchema.optional(),
    type: ShapeTypeSchema,
    text: z.string().max(2000).optional(),
  })
  .merge(ShapePlacementSchema)
  .merge(ShapeStyleSchema)
  .superRefine((v, ctx) => {
    const explicit = [v.left, v.top, v.width, v.height].every((n) => n !== undefined);
    if (!v.anchor && !explicit) {
      ctx.addIssue({ code: "custom", path: ["anchor"], message: "give `anchor` (cell range) or all of left/top/width/height" });
    }
  });
export type AddShapeInput = z.infer<typeof AddShapeInputSchema>;

/** Friendly dash names -> Excel line dash styles (see `DASH_STYLES`). */
export const LineDashSchema = z.enum(["solid", "dash", "longDash", "dashDot", "dashDotDot", "dot"]);
export type LineDash = z.infer<typeof LineDashSchema>;
export const DASH_STYLES: Record<LineDash, string> = {
  solid: "Solid",
  dash: "Dash",
  longDash: "LongDash",
  dashDot: "DashDot",
  dashDotDot: "DashDotDot",
  dot: "SquareDot",
};

export const ConnectorSideSchema = z.enum(["auto", "top", "bottom", "left", "right"]);
export type ConnectorSide = z.infer<typeof ConnectorSideSchema>;

export const AddConnectorInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  name: ShapeNameSchema.optional(),
  from: ShapeNameSchema,
  to: ShapeNameSchema,
  /** `auto` picks the facing sides from the two shapes' positions. */
  fromSide: ConnectorSideSchema.default("auto"),
  toSide: ConnectorSideSchema.default("auto"),
  /** Raw connection-site indexes; override the sides (needed for other shape types). */
  fromSite: z.number().int().min(0).max(63).optional(),
  toSite: z.number().int().min(0).max(63).optional(),
  connectorType: z.enum(["Straight", "Elbow", "Curve"]).default("Elbow"),
  arrow: z.enum(["end", "both", "none", "diamond"]).default("end").describe("`diamond` = line ending in a diamond (screen end)."),
  dash: LineDashSchema.default("solid"),
  color: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
  weight: z.number().min(0.25).max(20).optional(),
});
export type AddConnectorInput = z.infer<typeof AddConnectorInputSchema>;

export const UpdateShapeInputSchema = z
  .object({
    sheet: SheetRefSchema.optional(),
    name: ShapeNameSchema,
    newName: ShapeNameSchema.optional(),
    text: z.string().max(2000).optional(),
  })
  .merge(ShapePlacementSchema)
  .merge(ShapeStyleSchema);
export type UpdateShapeInput = z.infer<typeof UpdateShapeInputSchema>;

export const DeleteShapesInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  names: z.array(ShapeNameSchema).min(1).max(200),
});
export type DeleteShapesInput = z.infer<typeof DeleteShapesInputSchema>;

export const ListShapesInputSchema = z.object({ sheet: SheetRefSchema.optional() });
export type ListShapesInput = z.infer<typeof ListShapesInputSchema>;

export const InsertRowsInputSchema = z.object({
  sheet: SheetRefSchema.optional(),
  /** Row anchor, e.g. `5` or `A5`. Range insert is out of scope for v1. */
  at: z.string().regex(/^[A-Za-z]*\d+$/, "must be a row reference like 5 or A5"),
  count: z.number().int().min(1).max(1000),
  position: z.enum(["above", "below"]).default("above"),
});
export type InsertRowsInput = z.infer<typeof InsertRowsInputSchema>;

/** Fixed caps the Office.js executor applies on top of the bridge's tunable limits. */
export const EXECUTOR_CAPS = {
  /** Cells included in a selection snapshot. */
  selectionCells: 2_000,
  /** Cells returned by one read_range / search page. */
  readCells: 20_000,
  /** Cells scanned by one search_data call. */
  searchScanCells: 500_000,
  /** Non-empty cells named in a `would_overwrite` error. */
  overwritePreview: 20,
} as const;
