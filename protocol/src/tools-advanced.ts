import { z } from "zod";
import { A1AddressSchema, CellValueSchema, LineDashSchema, RangeRefSchema, SheetRefSchema } from "./tools.js";

const Hex = z.string().regex(/^#[0-9A-Fa-f]{6}$/, "must be #RRGGBB");
const CellAddress = z.string().regex(/^\$?[A-Za-z]{1,3}\$?\d{1,7}$/, "must be a single cell like B2");
const Letters = z.string().regex(/^[A-Za-z]{1,3}$/, "must be column letters like C");
const Sheet = SheetRefSchema.optional();

/* -------------------------------------------------------------------------- */
/* Cells: richer read / write / copy / resize / image                          */
/* -------------------------------------------------------------------------- */

export const SizeSchema = z
  .object({
    type: z.enum(["points", "standard", "autofit"]).describe("`standard` resets to the default size; `autofit` fits the content."),
    value: z.number().positive().max(1000).optional().describe("Size in points (required for type `points`)."),
  })
  .refine((v) => v.type !== "points" || v.value !== undefined, { message: "value is required for type points", path: ["value"] });
export type Size = z.infer<typeof SizeSchema>;

export const CellStyleSchema = z
  .object({
    fontColor: Hex.optional(),
    fontSize: z.number().min(1).max(409).optional(),
    fontFamily: z.string().max(64).optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strikethrough: z.boolean().optional(),
    backgroundColor: Hex.optional(),
    horizontalAlignment: z.enum(["Left", "Center", "Right", "General"]).optional(),
    wrapText: z.boolean().optional(),
    numberFormat: z.string().max(255).optional().describe('e.g. "@" text, "0.00", "$#,##0.00", "yyyy-mm-dd".'),
  })
  .strict();
export type CellStyle = z.infer<typeof CellStyleSchema>;

export const BorderSideSchema = z
  .object({
    style: z.enum(["solid", "dashed", "dotted", "double"]).default("solid"),
    weight: z.enum(["thin", "medium", "thick"]).default("thin"),
    color: Hex.default("#000000"),
  })
  .strict();
export type BorderSide = z.infer<typeof BorderSideSchema>;

export const BordersSchema = z
  .object({
    all: BorderSideSchema.optional().describe("Shortcut: same border on all four sides."),
    top: BorderSideSchema.optional(),
    bottom: BorderSideSchema.optional(),
    left: BorderSideSchema.optional(),
    right: BorderSideSchema.optional(),
  })
  .strict();
export type Borders = z.infer<typeof BordersSchema>;

export const CellSpecSchema = z
  .object({
    value: CellValueSchema.optional().describe("Static value. Put formulas in `formula`, not here."),
    formula: z.string().max(8192).optional().describe("Formula, e.g. =SUM(B2:B10)."),
    note: z.string().max(2000).optional().describe("Cell note (needs a recent Excel)."),
    style: CellStyleSchema.optional(),
    borders: BordersSchema.optional(),
  })
  .strict()
  .refine((c) => !(c.value !== undefined && c.formula !== undefined), { message: "give either value or formula, not both" });
export type CellSpec = z.infer<typeof CellSpecSchema>;

export const MAX_SET_CELLS = 2_000;

export const SetCellsInputSchema = z.object({
  sheet: Sheet,
  cells: z
    .record(CellAddress, CellSpecSchema)
    .refine((c) => Object.keys(c).length >= 1 && Object.keys(c).length <= MAX_SET_CELLS, { message: `1-${MAX_SET_CELLS} cells per call` })
    .describe('Map of cell address to cell data, e.g. {"A1": {"value": "Name", "style": {"bold": true}}, "B2": {"formula": "=A2*2"}}.'),
  copyTo: RangeRefSchema.optional().describe(
    "Copy the written cells (values, formulas AND formatting) to this larger range; the pattern repeats and relative references adjust. Counts as an overwrite of non-empty destination cells.",
  ),
  allowOverwrite: z.boolean().default(false).describe("Writing over non-empty cells fails unless true; set it only after the user confirmed."),
  resizeWidth: SizeSchema.optional().describe("Resize the columns of the written block afterwards."),
  resizeHeight: SizeSchema.optional().describe("Resize the rows of the written block afterwards."),
});
export type SetCellsInput = z.infer<typeof SetCellsInputSchema>;

export const GetCellRangesInputSchema = z.object({
  sheet: Sheet,
  ranges: z.array(RangeRefSchema).min(1).max(20).describe('A1 ranges, e.g. ["A1:C10", "E1:F5", "D:D", "5:5"].'),
  includeStyles: z.boolean().default(true).describe("Include fonts, fills, alignment, number formats and borders (grouped by identical style)."),
  cellLimit: z.number().int().min(1).max(20_000).default(2_000).describe("Max non-empty cells to return; the rest comes back in `nextRanges`."),
});
export type GetCellRangesInput = z.infer<typeof GetCellRangesInputSchema>;

export const FindErrorsInputSchema = z.object({
  sheet: Sheet.describe("Limit to one sheet; omit to scan every visible sheet."),
  address: A1AddressSchema.optional().describe("Limit the scan to this A1 range (needs `sheet`)."),
  limit: z.number().int().min(1).max(500).default(100).describe("Max error cells listed; totals still count every error found."),
});
export type FindErrorsInput = z.infer<typeof FindErrorsInputSchema>;

export const CheckFormulaConsistencyInputSchema = z.object({
  sheet: Sheet.describe("Limit to one sheet; omit to check every visible sheet."),
  address: A1AddressSchema.optional().describe("Limit the check to this A1 range (needs `sheet`); default is the used range."),
  direction: z
    .enum(["columns", "rows"])
    .default("columns")
    .describe("`columns`: a formula should repeat down each column (typical tables). `rows`: across each row (typical time series)."),
  limit: z.number().int().min(1).max(200).default(50).describe("Max issues listed; totals still count all of them."),
});
export type CheckFormulaConsistencyInput = z.infer<typeof CheckFormulaConsistencyInputSchema>;

export const MAX_REPLACE_CELLS = 2_000;

export const FindReplaceInputSchema = z.object({
  sheet: Sheet.describe("Limit to one sheet; omit for every visible sheet."),
  address: A1AddressSchema.optional().describe("Limit to this A1 range (needs `sheet`)."),
  find: z.string().min(1).max(500),
  replace: z.string().max(2000).describe("Replacement text; may be empty to delete the found text. Taken literally (no regex groups)."),
  matchCase: z.boolean().default(false),
  wholeCell: z.boolean().default(false).describe("Only cells whose entire text equals `find`."),
  inFormulas: z.boolean().default(false).describe("Also replace inside formula text. By default only text constants change; numbers and formulas are left alone."),
  apply: z.boolean().default(false).describe("false (default) only previews the changes; true makes them."),
  limit: z.number().int().min(1).max(500).default(100).describe("Max cells listed in the preview."),
});
export type FindReplaceInput = z.infer<typeof FindReplaceInputSchema>;

export const CopyToInputSchema = z.object({
  sheet: Sheet,
  source: RangeRefSchema,
  destination: RangeRefSchema.describe("May be larger than the source: the pattern repeats and formulas adjust (C2=B2 becomes C3=B3)."),
  copyType: z.enum(["all", "formulas", "values", "formats"]).default("all"),
  allowOverwrite: z.boolean().default(false),
});
export type CopyToInput = z.infer<typeof CopyToInputSchema>;

export const ResizeRangeInputSchema = z
  .object({
    sheet: Sheet,
    address: RangeRefSchema.optional().describe("Columns like A:D, rows like 1:5 or a cell range. Omit for the whole sheet."),
    width: SizeSchema.optional(),
    height: SizeSchema.optional(),
  })
  .refine((v) => v.width || v.height, { message: "give width and/or height", path: ["width"] });
export type ResizeRangeInput = z.infer<typeof ResizeRangeInputSchema>;

export const ReadRangeImageInputSchema = z.object({
  sheet: Sheet,
  address: A1AddressSchema.describe("Range to render, e.g. A1:F20. Keep under ~50 rows for legibility."),
});
export type ReadRangeImageInput = z.infer<typeof ReadRangeImageInputSchema>;

/* -------------------------------------------------------------------------- */
/* Sheets                                                                      */
/* -------------------------------------------------------------------------- */

export const ManageSheetInputSchema = z
  .object({
    operation: z.enum(["add", "delete", "rename", "duplicate", "move", "activate", "hide", "show", "set_tab_color"]),
    sheet: Sheet.describe("Target sheet (defaults to the active one; required for delete)."),
    name: z.string().min(1).max(31).optional().describe("Name for the new sheet (add / duplicate)."),
    newName: z.string().min(1).max(31).optional().describe("New name (rename)."),
    position: z.number().int().min(0).max(255).optional().describe("0-based tab position (add / duplicate / move)."),
    tabColor: z.union([Hex, z.literal("none")]).optional(),
    confirmDelete: z.boolean().default(false).describe("Deleting a sheet destroys its data: set true only after the user confirmed."),
  })
  .superRefine((v, ctx) => {
    const need = (ok: boolean, path: string, message: string) => ok || ctx.addIssue({ code: "custom", path: [path], message });
    if (v.operation === "rename") need(!!v.newName, "newName", "newName is required to rename");
    if (v.operation === "move") need(v.position !== undefined, "position", "position is required to move");
    if (v.operation === "set_tab_color") need(!!v.tabColor, "tabColor", "tabColor is required");
    if (v.operation === "delete") need(!!(v.sheet?.name || v.sheet?.id), "sheet", "name the sheet to delete");
    if (/[\\/?*[\]:]/.test(v.name ?? "") || /[\\/?*[\]:]/.test(v.newName ?? "")) {
      ctx.addIssue({ code: "custom", path: ["name"], message: "sheet names cannot contain \\ / ? * [ ] :" });
    }
  });
export type ManageSheetInput = z.infer<typeof ManageSheetInputSchema>;

/* -------------------------------------------------------------------------- */
/* Rules: conditional formatting, data validation                              */
/* -------------------------------------------------------------------------- */

const Operator = z.enum([
  "between",
  "notBetween",
  "equalTo",
  "notEqualTo",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
]);
export type RuleOperator = z.infer<typeof Operator>;

const CfFormat = z
  .object({
    fillColor: Hex.optional(),
    fontColor: Hex.optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    strikethrough: z.boolean().optional(),
  })
  .strict()
  .refine((f) => Object.keys(f).length > 0, { message: "give at least one format property" });
export type CfFormat = z.infer<typeof CfFormat>;

export const ICON_SETS = [
  "threeArrows",
  "threeArrowsGray",
  "threeTriangles",
  "threeTrafficLights1",
  "threeTrafficLights2",
  "threeSigns",
  "threeSymbols",
  "threeFlags",
  "fourArrows",
  "fourTrafficLights",
  "fiveArrows",
  "fiveRating",
] as const;

export const CfRuleSchema = z.discriminatedUnion("type", [
  z.object({ type: z.literal("cellValue"), operator: Operator, value: z.union([z.string(), z.number()]), value2: z.union([z.string(), z.number()]).optional(), format: CfFormat }),
  z.object({ type: z.literal("containsText"), text: z.string().min(1), match: z.enum(["contains", "notContains", "beginsWith", "endsWith"]).default("contains"), format: CfFormat }),
  z.object({ type: z.literal("custom"), formula: z.string().min(1).describe('Formula relative to the top-left cell, e.g. =$B2>100.'), format: CfFormat }),
  z.object({ type: z.literal("colorScale"), minColor: Hex, midColor: Hex.optional(), maxColor: Hex }),
  z.object({ type: z.literal("dataBar"), color: Hex.optional(), negativeColor: Hex.optional(), showValue: z.boolean().default(true) }),
  z.object({ type: z.literal("iconSet"), style: z.enum(ICON_SETS), reverse: z.boolean().default(false), showValue: z.boolean().default(true) }),
  z.object({ type: z.literal("topBottom"), kind: z.enum(["topItems", "topPercent", "bottomItems", "bottomPercent"]), rank: z.number().int().min(1).max(1000), format: CfFormat }),
  z.object({ type: z.literal("duplicates"), unique: z.boolean().default(false).describe("true highlights unique values instead."), format: CfFormat }),
]);
export type CfRule = z.infer<typeof CfRuleSchema>;

export const ConditionalFormatInputSchema = z
  .object({
    sheet: Sheet,
    operation: z.enum(["add", "list", "clear"]),
    address: RangeRefSchema.optional().describe("Range the rule applies to (required for add / clear)."),
    rule: CfRuleSchema.optional().describe("The rule to add."),
  })
  .superRefine((v, ctx) => {
    if (v.operation === "add" && (!v.address || !v.rule)) ctx.addIssue({ code: "custom", path: ["rule"], message: "add needs address and rule" });
    if (v.operation === "clear" && !v.address) ctx.addIssue({ code: "custom", path: ["address"], message: "clear needs address" });
  });
export type ConditionalFormatInput = z.infer<typeof ConditionalFormatInputSchema>;

const Bound = z.union([z.string(), z.number()]);

export const ValidationRuleSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("list"),
    source: z.union([z.array(z.string().max(255)).min(1).max(200), z.string().startsWith("=")]).describe('Either the allowed items, or a range formula such as "=$A$2:$A$9" / "=Sheet2!$A$1:$A$5".'),
    inCellDropDown: z.boolean().default(true),
  }),
  z.object({ type: z.enum(["whole", "decimal", "date", "time", "textLength"]), operator: Operator, value1: Bound.describe("Number, ISO date, or a formula starting with =."), value2: Bound.optional() }),
  z.object({ type: z.literal("custom"), formula: z.string().startsWith("=") }),
]);
export type ValidationRule = z.infer<typeof ValidationRuleSchema>;

export const DataValidationInputSchema = z
  .object({
    sheet: Sheet,
    address: RangeRefSchema,
    operation: z.enum(["set", "get", "clear"]),
    rule: ValidationRuleSchema.optional(),
    ignoreBlanks: z.boolean().default(true),
    prompt: z.object({ title: z.string().max(32).optional(), message: z.string().max(255) }).optional().describe("Input hint shown when the cell is selected."),
    errorAlert: z
      .object({ style: z.enum(["stop", "warning", "information"]).default("stop"), title: z.string().max(32).optional(), message: z.string().max(255) })
      .optional(),
  })
  .superRefine((v, ctx) => {
    if (v.operation === "set" && !v.rule) ctx.addIssue({ code: "custom", path: ["rule"], message: "set needs a rule" });
    if (v.rule && v.rule.type !== "list" && v.rule.type !== "custom" && ["between", "notBetween"].includes(v.rule.operator) && v.rule.value2 === undefined) {
      ctx.addIssue({ code: "custom", path: ["rule", "value2"], message: "between/notBetween need value2" });
    }
  });
export type DataValidationInput = z.infer<typeof DataValidationInputSchema>;

/* -------------------------------------------------------------------------- */
/* Sort / filter / merge / names / page layout                                 */
/* -------------------------------------------------------------------------- */

export const SortRangeInputSchema = z.object({
  sheet: Sheet,
  address: A1AddressSchema.describe("Data range to sort (include the header row when hasHeaders is true)."),
  keys: z
    .array(z.object({ column: Letters.describe("Sheet column letter, e.g. C"), ascending: z.boolean().default(true) }))
    .min(1)
    .max(8),
  hasHeaders: z.boolean().default(true),
  matchCase: z.boolean().default(false),
});
export type SortRangeInput = z.infer<typeof SortRangeInputSchema>;

const FilterOperator = z.enum([
  "equals",
  "notEquals",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
  "contains",
  "notContains",
  "beginsWith",
  "endsWith",
  "blank",
  "nonBlank",
]);
export type FilterOperator = z.infer<typeof FilterOperator>;

export const FilterCriteriaSchema = z.union([
  z.object({ values: z.array(z.string()).min(1).max(500).describe("Show only rows whose cell equals one of these.") }),
  z.object({
    conditions: z.array(z.object({ operator: FilterOperator, value: Bound.optional() })).min(1).max(2),
    join: z.enum(["and", "or"]).default("and"),
  }),
  z.object({ top: z.object({ kind: z.enum(["topItems", "topPercent", "bottomItems", "bottomPercent"]), n: z.number().int().min(1).max(1000) }) }),
]);
export type FilterCriteria = z.infer<typeof FilterCriteriaSchema>;

export const FilterRangeInputSchema = z
  .object({
    sheet: Sheet,
    operation: z.enum(["apply", "clear", "reapply", "remove"]).describe("apply: turn on the filter (optionally filtering a column); clear: show all rows; remove: turn the filter off."),
    address: A1AddressSchema.optional().describe("Data range including headers (required for apply)."),
    column: Letters.optional().describe("Sheet column letter to filter on (needs criteria)."),
    criteria: FilterCriteriaSchema.optional(),
  })
  .superRefine((v, ctx) => {
    if (v.operation === "apply" && !v.address) ctx.addIssue({ code: "custom", path: ["address"], message: "apply needs address" });
    if (!!v.criteria !== !!v.column) ctx.addIssue({ code: "custom", path: ["column"], message: "column and criteria go together" });
  });
export type FilterRangeInput = z.infer<typeof FilterRangeInputSchema>;

export const MergeRangeInputSchema = z.object({
  sheet: Sheet,
  address: A1AddressSchema,
  operation: z.enum(["merge", "unmerge"]).default("merge"),
  across: z.boolean().default(false).describe("Merge each row separately."),
  allowOverwrite: z.boolean().default(false).describe("Merging keeps only the top-left value; other non-empty cells need this."),
});
export type MergeRangeInput = z.infer<typeof MergeRangeInputSchema>;

export const ManageNamesInputSchema = z
  .object({
    operation: z.enum(["add", "list", "delete"]),
    name: z.string().min(1).max(255).regex(/^[A-Za-z_\\][A-Za-z0-9_.]*$/, "invalid name").optional(),
    reference: z.string().max(1024).optional().describe('Range like "Sheet1!$A$1:$A$5" or a formula/constant like "=0.07".'),
    scope: z.string().max(64).default("workbook").describe('`workbook` or a sheet name.'),
    comment: z.string().max(255).optional(),
  })
  .superRefine((v, ctx) => {
    if (v.operation !== "list" && !v.name) ctx.addIssue({ code: "custom", path: ["name"], message: "name is required" });
    if (v.operation === "add" && !v.reference) ctx.addIssue({ code: "custom", path: ["reference"], message: "reference is required" });
  });
export type ManageNamesInput = z.infer<typeof ManageNamesInputSchema>;

export const PageLayoutInputSchema = z.object({
  sheet: Sheet,
  printArea: A1AddressSchema.optional().describe("Range to print, e.g. A1:H40."),
  orientation: z.enum(["portrait", "landscape"]).optional(),
  paperSize: z.enum(["Letter", "Legal", "A3", "A4", "A5", "Tabloid"]).optional(),
  fitToPagesWide: z.number().int().min(1).max(100).optional(),
  fitToPagesTall: z.number().int().min(0).max(100).optional().describe("0 = as many pages as needed."),
  scalePercent: z.number().int().min(10).max(400).optional().describe("Alternative to fitToPages*."),
  repeatTitleRows: z.string().regex(/^\d+:\d+$/).optional().describe('e.g. "1:1"'),
  repeatTitleColumns: z.string().regex(/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/).optional().describe('e.g. "A:A"'),
  marginsInches: z.object({ left: z.number().min(0).max(10), right: z.number().min(0).max(10), top: z.number().min(0).max(10), bottom: z.number().min(0).max(10) }).partial().optional(),
  header: z.object({ left: z.string().optional(), center: z.string().optional(), right: z.string().optional() }).optional(),
  footer: z.object({ left: z.string().optional(), center: z.string().optional(), right: z.string().optional() }).optional(),
  centerHorizontally: z.boolean().optional(),
  centerVertically: z.boolean().optional(),
  printGridlines: z.boolean().optional(),
  printHeadings: z.boolean().optional(),
});
export type PageLayoutInput = z.infer<typeof PageLayoutInputSchema>;

/* -------------------------------------------------------------------------- */
/* Objects: charts, pivot tables, Excel Tables                                 */
/* -------------------------------------------------------------------------- */

/** Friendly chart names -> Excel.ChartType values. */
export const CHART_TYPES = {
  columnClustered: "ColumnClustered",
  columnStacked: "ColumnStacked",
  columnStacked100: "ColumnStacked100",
  barClustered: "BarClustered",
  barStacked: "BarStacked",
  barStacked100: "BarStacked100",
  line: "Line",
  lineMarkers: "LineMarkers",
  lineStacked: "LineStacked",
  area: "Area",
  areaStacked: "AreaStacked",
  areaStacked100: "AreaStacked100",
  pie: "Pie",
  pieExploded: "PieExploded",
  doughnut: "Doughnut",
  scatter: "XYScatter",
  scatterLines: "XYScatterLinesNoMarkers",
  scatterLinesMarkers: "XYScatterLines",
  scatterSmooth: "XYScatterSmooth",
  bubble: "Bubble",
  radar: "Radar",
  radarMarkers: "RadarMarkers",
  radarFilled: "RadarFilled",
  waterfall: "Waterfall",
  histogram: "Histogram",
  treemap: "Treemap",
  sunburst: "Sunburst",
  funnel: "Funnel",
  boxWhisker: "BoxwhiskerChart",
  stockHLC: "StockHLC",
  stockOHLC: "StockOHLC",
} as const;
export type ChartTypeName = keyof typeof CHART_TYPES;
const ChartTypeSchema = z.enum(Object.keys(CHART_TYPES) as [ChartTypeName, ...ChartTypeName[]]);

const ObjectName = z.string().min(1).max(255);
const OperationSchema = z.enum(["create", "update", "delete"]);
const needsName = (v: { operation: string; name?: string }, ctx: z.RefinementCtx) => {
  if (v.operation !== "create" && !v.name) ctx.addIssue({ code: "custom", path: ["name"], message: "name is required for update/delete" });
};

export const GetObjectsInputSchema = z.object({
  sheet: Sheet.describe("Limit to one sheet; omit for the whole workbook."),
  name: ObjectName.optional().describe("Return only the object with this name."),
});
export type GetObjectsInput = z.infer<typeof GetObjectsInputSchema>;

export const ModifyChartInputSchema = z
  .object({
    operation: OperationSchema,
    sheet: Sheet,
    name: ObjectName.optional().describe("Chart name (from get_objects). Optional on create."),
    newName: ObjectName.optional(),
    chartType: ChartTypeSchema.optional(),
    title: z.string().max(255).optional(),
    source: z.string().max(256).optional().describe("Contiguous data range such as Sheet1!A1:D10: headers in the first row become series names, the first column the categories."),
    seriesBy: z.enum(["auto", "rows", "columns"]).default("auto"),
    anchor: CellAddress.optional().describe("Top-left cell the chart sits on, e.g. F2."),
    width: z.number().min(50).max(5000).optional().describe("Points."),
    height: z.number().min(50).max(5000).optional().describe("Points."),
    legend: z.enum(["top", "bottom", "left", "right", "none"]).optional(),
    xAxisTitle: z.string().max(255).optional(),
    yAxisTitle: z.string().max(255).optional(),
  })
  .superRefine((v, ctx) => {
    needsName(v, ctx);
    if (v.operation === "create" && (!v.chartType || !v.source)) ctx.addIssue({ code: "custom", path: ["source"], message: "create needs chartType and source" });
  });
export type ModifyChartInput = z.infer<typeof ModifyChartInputSchema>;

export const SUMMARIZE_BY = ["sum", "count", "average", "max", "min", "product", "countNumbers", "standardDeviation", "variance"] as const;

export const ModifyPivotTableInputSchema = z
  .object({
    operation: OperationSchema,
    sheet: Sheet.describe("Sheet the pivot table lives on (destination sheet for create)."),
    name: ObjectName.describe("Pivot table name."),
    newName: ObjectName.optional(),
    source: z.string().max(256).optional().describe("Source data incl. headers, e.g. Sheet1!A1:D100. Fixed after creation."),
    destination: CellAddress.optional().describe("Top-left cell for the pivot table, e.g. F1. Fixed after creation."),
    rows: z.array(z.string()).optional().describe("Field (header) names to group by vertically."),
    columns: z.array(z.string()).optional(),
    filters: z.array(z.string()).optional(),
    values: z
      .array(z.object({ field: z.string(), summarizeBy: z.enum(SUMMARIZE_BY).default("sum"), label: z.string().max(255).optional() }))
      .optional()
      .describe("Fields to aggregate (at least one on create)."),
  })
  .superRefine((v, ctx) => {
    if (v.operation === "create") {
      if (!v.source || !v.destination) ctx.addIssue({ code: "custom", path: ["source"], message: "create needs source and destination" });
      if (!v.values?.length) ctx.addIssue({ code: "custom", path: ["values"], message: "create needs at least one value field" });
    }
    if (v.operation === "update" && (v.source || v.destination)) {
      ctx.addIssue({ code: "custom", path: ["source"], message: "source/destination cannot change; delete and recreate the pivot table" });
    }
  });
export type ModifyPivotTableInput = z.infer<typeof ModifyPivotTableInputSchema>;

export const ModifyTableInputSchema = z
  .object({
    operation: OperationSchema,
    sheet: Sheet,
    name: ObjectName.optional().describe("Table name. Optional on create."),
    newName: ObjectName.optional(),
    range: A1AddressSchema.optional().describe("Data range for create, e.g. A1:D20."),
    hasHeaders: z.boolean().default(true),
    resizeTo: A1AddressSchema.optional().describe("New data range for update."),
    style: z.string().regex(/^TableStyle(Light|Medium|Dark)\d{1,2}$/).optional().describe("e.g. TableStyleMedium2"),
    showTotals: z.boolean().optional(),
    bandedRows: z.boolean().optional(),
    showFilterButton: z.boolean().optional(),
    deleteData: z.boolean().default(false).describe("delete: also erase the cells; default keeps the data and only removes the table."),
  })
  .superRefine((v, ctx) => {
    needsName(v, ctx);
    if (v.operation === "create" && !v.range) ctx.addIssue({ code: "custom", path: ["range"], message: "create needs range" });
  });
export type ModifyTableInput = z.infer<typeof ModifyTableInputSchema>;

/* -------------------------------------------------------------------------- */
/* Escape hatch (opt-in)                                                       */
/* -------------------------------------------------------------------------- */

export const ExecuteOfficeJsInputSchema = z.object({
  code: z
    .string()
    .min(1)
    .max(20_000)
    .describe(
      "Body of an async function with `context` (Excel.RequestContext), `Excel` and `Office` in scope, run inside Excel.run. " +
        "`return` a JSON-serialisable value. Call `await context.sync()` yourself.",
    ),
  description: z.string().max(200).optional().describe("One line saying what this does (shown in the change list)."),
});
export type ExecuteOfficeJsInput = z.infer<typeof ExecuteOfficeJsInputSchema>;

/* -------------------------------------------------------------------------- */
/* Free-standing line (legends, simple arrows)                                 */
/* -------------------------------------------------------------------------- */

export const AddLineInputSchema = z.object({
  sheet: Sheet,
  name: z.string().min(1).max(64).optional(),
  from: CellAddress.describe("Start cell; the line starts at its centre, e.g. D7."),
  to: CellAddress.describe("End cell; the line ends at its centre, e.g. G7."),
  arrow: z.enum(["end", "both", "none", "diamond"]).default("end"),
  dash: LineDashSchema.default("solid"),
  color: Hex.optional(),
  weight: z.number().min(0.25).max(20).optional(),
});
export type AddLineInput = z.infer<typeof AddLineInputSchema>;

/* -------------------------------------------------------------------------- */
/* Skills (answered by the bridge, never sent to the task pane)                */
/* -------------------------------------------------------------------------- */

export const ListSkillsInputSchema = z.object({});
export const SaveSkillInputSchema = z.object({
  name: z.string().regex(/^[a-z0-9][a-z0-9_-]{0,63}$/).describe("New skill name, used as its folder name and slash command."),
  description: z.string().max(2000).default("").describe("Short summary of when to use this skill."),
  instructions: z.string().min(1).max(190_000).describe("Markdown instructions only, without YAML frontmatter. Saved as SKILL.md in the user's Excel Partner skills folder."),
}).strict();
export type SaveSkillInput = z.infer<typeof SaveSkillInputSchema>;
export const UpdateSkillInputSchema = z.object({
  name: z.string().regex(/^[a-z0-9][a-z0-9_-]{0,63}$/),
  file: z.string().min(1).max(200).default("SKILL.md"),
  content: z.string().max(5 * 1024 * 1024),
  revision: z.string().regex(/^[a-f0-9]{64}$/).describe("Revision returned by read_skill with forEditing=true. Prevents overwriting newer edits."),
}).strict();
export type UpdateSkillInput = z.infer<typeof UpdateSkillInputSchema>;
export const ReadSkillInputSchema = z.object({
  name: z.string().min(1).max(64).describe("Skill name from list_skills."),
  file: z.string().min(1).max(200).optional().describe("Reference file listed by a previous read_skill, e.g. references/diagram-style.md."),
  forEditing: z.boolean().optional().describe("Return the exact UTF-8 file and its revision for update_skill; use this before changing any installed skill file."),
});
export type ReadSkillInput = z.infer<typeof ReadSkillInputSchema>;

export const RunSkillScriptInputSchema = z.object({
  name: z.string().regex(/^[a-z0-9][a-z0-9_-]{0,63}$/).describe("Installed skill name."),
  file: z.string().min(1).max(200).describe("Python (.py) or Bash (.sh) script relative to the skill folder, e.g. tools/flow_layout.py."),
  args: z.array(z.string().max(4096).refine(value => !value.includes("\0"), "Arguments cannot contain NUL")).max(64).default([])
    .describe("Literal arguments, without shell quoting. Working directory is the skill folder."),
  stdin: z.string().max(200_000).optional().describe("Optional text supplied on standard input, such as JSON."),
  outputFiles: z.array(z.string().min(1).max(200)).max(8).default([])
    .describe("Optional generated files relative to the skill folder to return as attachment:// paths (30 MB combined)."),
  timeoutMs: z.number().int().min(100).max(60_000).default(30_000),
}).strict();
export type RunSkillScriptInput = z.infer<typeof RunSkillScriptInputSchema>;

const PythonWorkspacePath = z.string().min(1).max(200).refine(value =>
  value.split("/").every(part => Boolean(part) && part !== "." && part !== ".." && !/[\\:\x00-\x1f]/.test(part)),
"Use a relative path without traversal or drive letters.");
export const RunPythonInputSchema = z.object({
  code: z.string().min(1).max(190_000).describe("Python source to write as main.py and execute with Python installed on the user's machine. Local user permissions and installed libraries apply."),
  args: z.array(z.string().max(4096).refine(value => !value.includes("\0"))).max(64).default([]),
  stdin: z.string().max(200_000).default(""),
  files: z.array(z.object({
    path: PythonWorkspacePath.refine(value => value !== "main.py", "main.py is reserved for code"),
    attachment: z.string().regex(/^attachment:\/\/[0-9a-f-]{36}$/).describe("Existing attachment:// reference to copy into a temporary Python working folder."),
  }).strict()).max(16).default([]),
  outputFiles: z.array(PythonWorkspacePath).max(8).default([]).describe("Generated files to return as attachments; maximum 10 MB combined."),
  timeoutMs: z.number().int().min(100).max(60_000).default(30_000),
}).strict();
export type RunPythonInput = z.infer<typeof RunPythonInputSchema>;

/* -------------------------------------------------------------------------- */
/* Reference files (read-only access to folders the user opened up)            */
/* -------------------------------------------------------------------------- */

export const ListReferenceFilesInputSchema = z.object({
  path: z.string().min(1).max(500).optional().describe("Folder to list (Windows or WSL path). Omit to see the folders you are allowed to read."),
});
export type ListReferenceFilesInput = z.infer<typeof ListReferenceFilesInputSchema>;

export const ReadWorkbookFileInputSchema = z.object({
  path: z.string().min(1).max(500).describe("Path of a closed .xlsx, e.g. C:\\Users\\me\\Template.xlsx."),
  sheet: z.string().min(1).max(255).optional().describe("Sheet to inspect; omit for the workbook overview (sheets, print areas)."),
  range: A1AddressSchema.optional().describe("Limit to a range such as A1:BC30 (default: the used range)."),
  includeStyles: z.boolean().default(true).describe("Fonts, fills, alignment and borders grouped by identical style."),
  includeShapes: z.boolean().default(true).describe("Drawn shapes and connectors with geometry, text, fill, line and anchors."),
  maxCells: z.number().int().min(1).max(10_000).default(3_000),
});
export type ReadWorkbookFileInput = z.infer<typeof ReadWorkbookFileInputSchema>;

export const ReadAttachedImageInputSchema = z.object({
  path: z.string().min(1).max(500)
    .describe("attachment:// path, selected local image path, or skill-template:// path listed by read_skill."),
});
export type ReadAttachedImageInput = z.infer<typeof ReadAttachedImageInputSchema>;

export const ReadAttachedFileInputSchema = z.object({
  path: z.string().min(1).max(500).describe("attachment:// path or selected local file path."),
  entry: z.string().min(1).max(500).optional().describe("For ZIP files, the entry to read. Omit to list entries."),
  offset: z.number().int().min(0).default(0).describe("Byte offset for large files."),
  maxBytes: z.number().int().min(1).max(65536).default(16384),
});
export type ReadAttachedFileInput = z.infer<typeof ReadAttachedFileInputSchema>;

export const InsertSheetsFromFileInputSchema = z.object({
  path: z.string().min(1).max(500).describe("Closed .xlsx whose sheets are copied into the OPEN workbook (formatting, merges, shapes included)."),
  sheets: z.array(z.string().min(1).max(255)).min(1).max(30).optional().describe("Sheet names to copy; omit for all."),
  position: z.enum(["end", "beginning", "after"]).default("end"),
  afterSheet: z.string().min(1).max(255).optional().describe("For position `after`: the open workbook's sheet to insert after."),
});
export type InsertSheetsFromFileInput = z.infer<typeof InsertSheetsFromFileInputSchema>;

/* -------------------------------------------------------------------------- */
/* Artifacts (files the user previews in the task pane)                        */
/* -------------------------------------------------------------------------- */

export const ARTIFACT_EXTENSIONS = ["html", "htm", "svg", "md", "markdown", "txt", "csv", "json"] as const;
export const MAX_ARTIFACT_CHARS = 300_000;

export const CreateArtifactInputSchema = z.object({
  filename: z.string().min(1).max(100)
    .regex(/^[\w][\w .-]*\.(html?|svg|md|markdown|txt|csv|json)$/i, `A plain file name ending in ${ARTIFACT_EXTENSIONS.map(e => "." + e).join(", ")}; no folders.`)
    .describe("File name with extension, e.g. report.html or notes.md."),
  content: z.string().min(1).max(MAX_ARTIFACT_CHARS).describe("Complete file text. Not a diff: to change a file, send the whole file again."),
}).strict();
export type CreateArtifactInput = z.infer<typeof CreateArtifactInputSchema>;
