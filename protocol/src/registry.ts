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
  ValidateFlowchartInputSchema,
  UpdateShapeInputSchema,
} from "./tools.js";
import {
  AddLineInputSchema,
  CheckFormulaConsistencyInputSchema,
  CreateArtifactInputSchema,
  FindReplaceInputSchema,
  ConditionalFormatInputSchema,
  CopyToInputSchema,
  DataValidationInputSchema,
  ExecuteOfficeJsInputSchema,
  FilterRangeInputSchema,
  FindErrorsInputSchema,
  GetCellRangesInputSchema,
  GetObjectsInputSchema,
  InsertSheetsFromFileInputSchema,
  ListReferenceFilesInputSchema,
  ListSkillsInputSchema,
  SaveSkillInputSchema,
  UpdateSkillInputSchema,
  ReadWorkbookFileInputSchema,
  ReadAttachedImageInputSchema,
  ReadAttachedFileInputSchema,
  ReadSkillInputSchema,
  RunSkillScriptInputSchema,
  RunPythonInputSchema,
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
import {
  DeleteSiteKnowledgeInputSchema,
  ListTestSessionsInputSchema,
  OpenTestBrowserInputSchema,
  ReadSiteKnowledgeInputSchema,
  RecordTestResultsInputSchema,
  SaveSiteKnowledgeInputSchema,
  SaveTestScriptInputSchema,
  FinishTestRunInputSchema,
} from "./tools-testing.js";

/** Every tool the Excel MCP server can expose, with its input schema. */
export const ToolInputSchemas = {
  // read
  get_selection: GetSelectionInputSchema,
  get_sheets: GetSheetsInputSchema,
  read_range: ReadRangeInputSchema,
  get_cell_ranges: GetCellRangesInputSchema,
  search_data: SearchDataInputSchema,
  find_errors: FindErrorsInputSchema,
  check_formula_consistency: CheckFormulaConsistencyInputSchema,
  read_range_image: ReadRangeImageInputSchema,
  get_objects: GetObjectsInputSchema,
  list_shapes: ListShapesInputSchema,
  validate_flowchart: ValidateFlowchartInputSchema,
  // cells
  write_range: WriteRangeInputSchema,
  set_cells: SetCellsInputSchema,
  set_formula: SetFormulaInputSchema,
  find_replace: FindReplaceInputSchema,
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
  update_skill: UpdateSkillInputSchema,
  read_skill: ReadSkillInputSchema,
  run_skill_script: RunSkillScriptInputSchema,
  run_python: RunPythonInputSchema,
  // shown to the user as a preview card in the task pane
  create_artifact: CreateArtifactInputSchema,
  // reference files: read-only, inside folders the user configured
  list_reference_files: ListReferenceFilesInputSchema,
  read_workbook_file: ReadWorkbookFileInputSchema,
  read_attached_image: ReadAttachedImageInputSchema,
  read_attached_file: ReadAttachedFileInputSchema,
  insert_sheets_from_file: InsertSheetsFromFileInputSchema,
  playwright_browser: z.object({ tool: z.string().regex(/^(list_tools|browser_[a-z_]+)$/), arguments: z.record(z.unknown()).default({}) }),
  // testing from a sheet: the browser is opened by the bridge, results are written by the task pane
  list_test_sessions: ListTestSessionsInputSchema,
  read_site_knowledge: ReadSiteKnowledgeInputSchema,
  save_site_knowledge: SaveSiteKnowledgeInputSchema,
  delete_site_knowledge: DeleteSiteKnowledgeInputSchema,
  open_test_browser: OpenTestBrowserInputSchema,
  record_test_results: RecordTestResultsInputSchema,
  save_test_script: SaveTestScriptInputSchema,
  finish_test_run: FinishTestRunInputSchema,
  // opt-in escape hatch
  execute_office_js: ExecuteOfficeJsInputSchema,
} as const satisfies Record<string, z.ZodTypeAny>;

export type ExcelToolName = keyof typeof ToolInputSchemas;
export const EXCEL_TOOL_NAMES = Object.keys(ToolInputSchemas) as ExcelToolName[];

/** Tools the bridge answers itself instead of forwarding to the task pane. */
export const SERVER_TOOLS: readonly ExcelToolName[] = ["playwright_browser", "list_skills", "read_skill", "save_skill", "update_skill", "run_skill_script", "run_python", "list_reference_files", "read_workbook_file", "read_attached_image", "read_attached_file", "list_test_sessions", "open_test_browser", "read_site_knowledge", "save_site_knowledge", "delete_site_knowledge", "save_test_script", "finish_test_run"];

/** Tools that need the user to have switched on the browser (Playwright) in Tools. */
export const BROWSER_TOOLS: readonly ExcelToolName[] = ["playwright_browser", "list_test_sessions", "open_test_browser", "save_test_script", "finish_test_run"];

/** Tools controlled by allowOfficeJs (enabled by default; BRIDGE_ALLOW_OFFICE_JS=0 disables them). */
export const OPT_IN_TOOLS: readonly ExcelToolName[] = ["execute_office_js"];

/** Tool names a session exposes, given whether the opt-in tools are enabled. */
export function exposedToolNames(allowOptIn: boolean, capabilities: { webSearch?: boolean; playwright?: boolean } = {}): ExcelToolName[] {
  return EXCEL_TOOL_NAMES.filter((name) => (allowOptIn || !OPT_IN_TOOLS.includes(name))
    && (!BROWSER_TOOLS.includes(name) || capabilities.playwright));
}

export function isExcelToolName(name: string): name is ExcelToolName {
  return (EXCEL_TOOL_NAMES as readonly string[]).includes(name);
}

/** Parse and validate tool arguments, returning a typed result or a Zod error. */
export function parseToolInput<T extends ExcelToolName>(tool: T, args: unknown): z.infer<(typeof ToolInputSchemas)[T]> {
  return (ToolInputSchemas[tool] as z.ZodTypeAny).parse(args ?? {}) as z.infer<(typeof ToolInputSchemas)[T]>;
}
