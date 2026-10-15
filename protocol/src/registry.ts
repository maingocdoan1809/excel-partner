import { z } from "zod";
import {
  ClearRangeInputSchema,
  FormatRangeInputSchema,
  GetSelectionInputSchema,
  GetSheetsInputSchema,
  InsertRowsInputSchema,
  ModifyStructureInputSchema,
  ReadRangeInputSchema,
  SearchDataInputSchema,
  SetFormulaInputSchema,
  WriteRangeInputSchema,
  AddShapeInputSchema,
  AddConnectorInputSchema,
  DeleteShapesInputSchema,
  ListShapesInputSchema,
  UpdateShapeInputSchema,
} from "./tools.js";
import {
  AddLineInputSchema,
  ConditionalFormatInputSchema,
  CopyToInputSchema,
  DataValidationInputSchema,
  ExecuteOfficeJsInputSchema,
  FilterRangeInputSchema,
  GetCellRangesInputSchema,
  GetObjectsInputSchema,
  InsertSheetsFromFileInputSchema,
  ListReferenceFilesInputSchema,
  ListSkillsInputSchema,
  SaveSkillInputSchema,
  ReadWorkbookFileInputSchema,
  ReadAttachedImageInputSchema,
  ReadAttachedFileInputSchema,
  ReadSkillInputSchema,
  ManageNamesInputSchema,
  ManageSheetInputSchema,
  MergeRangeInputSchema,
  ModifyChartInputSchema,
  ModifyPivotTableInputSchema,
  ModifyTableInputSchema,
  PageLayoutInputSchema,
  ReadRangeImageInputSchema,
  ResizeRangeInputSchema,
  SetCellsInputSchema,
  SortRangeInputSchema,
} from "./tools-advanced.js";

/** Every tool the Excel MCP server can expose, with its input schema. */
export const ToolInputSchemas = {
  // read
  get_selection: GetSelectionInputSchema,
  get_sheets: GetSheetsInputSchema,
  read_range: ReadRangeInputSchema,
  get_cell_ranges: GetCellRangesInputSchema,
  search_data: SearchDataInputSchema,
  read_range_image: ReadRangeImageInputSchema,
  get_objects: GetObjectsInputSchema,
  list_shapes: ListShapesInputSchema,
  // cells
  write_range: WriteRangeInputSchema,
  set_cells: SetCellsInputSchema,
  set_formula: SetFormulaInputSchema,
  copy_to: CopyToInputSchema,
  format_range: FormatRangeInputSchema,
  clear_range: ClearRangeInputSchema,
  resize_range: ResizeRangeInputSchema,
  merge_range: MergeRangeInputSchema,
  sort_range: SortRangeInputSchema,
  filter_range: FilterRangeInputSchema,
  conditional_format: ConditionalFormatInputSchema,
  data_validation: DataValidationInputSchema,
  // structure
  insert_rows: InsertRowsInputSchema,
  modify_structure: ModifyStructureInputSchema,
  manage_sheet: ManageSheetInputSchema,
  manage_names: ManageNamesInputSchema,
  page_layout: PageLayoutInputSchema,
  // objects
  modify_chart: ModifyChartInputSchema,
  modify_pivot_table: ModifyPivotTableInputSchema,
  modify_table: ModifyTableInputSchema,
  // drawing
  add_shape: AddShapeInputSchema,
  add_connector: AddConnectorInputSchema,
  add_line: AddLineInputSchema,
  update_shape: UpdateShapeInputSchema,
  delete_shapes: DeleteShapesInputSchema,
  // skills (answered by the bridge itself)
  list_skills: ListSkillsInputSchema,
  save_skill: SaveSkillInputSchema,
  read_skill: ReadSkillInputSchema,
  // reference files: read-only, inside folders the user configured
  list_reference_files: ListReferenceFilesInputSchema,
  read_workbook_file: ReadWorkbookFileInputSchema,
  read_attached_image: ReadAttachedImageInputSchema,
  read_attached_file: ReadAttachedFileInputSchema,
  insert_sheets_from_file: InsertSheetsFromFileInputSchema,
  // opt-in escape hatch
  execute_office_js: ExecuteOfficeJsInputSchema,
} as const satisfies Record<string, z.ZodTypeAny>;

export type ExcelToolName = keyof typeof ToolInputSchemas;
export const EXCEL_TOOL_NAMES = Object.keys(ToolInputSchemas) as ExcelToolName[];

/** Tools the bridge answers itself instead of forwarding to the task pane. */
export const SERVER_TOOLS: readonly ExcelToolName[] = ["list_skills", "read_skill", "save_skill", "list_reference_files", "read_workbook_file", "read_attached_image", "read_attached_file"];

/** Tools that are off unless the bridge is started with BRIDGE_ALLOW_OFFICE_JS=1. */
export const OPT_IN_TOOLS: readonly ExcelToolName[] = ["execute_office_js"];

/** Tool names a session exposes, given whether the opt-in tools are enabled. */
export function exposedToolNames(allowOptIn: boolean): ExcelToolName[] {
  return EXCEL_TOOL_NAMES.filter((name) => allowOptIn || !OPT_IN_TOOLS.includes(name));
}

export function isExcelToolName(name: string): name is ExcelToolName {
  return (EXCEL_TOOL_NAMES as readonly string[]).includes(name);
}

/** Parse and validate tool arguments, returning a typed result or a Zod error. */
export function parseToolInput<T extends ExcelToolName>(tool: T, args: unknown): z.infer<(typeof ToolInputSchemas)[T]> {
  return (ToolInputSchemas[tool] as z.ZodTypeAny).parse(args ?? {}) as z.infer<(typeof ToolInputSchemas)[T]>;
}
