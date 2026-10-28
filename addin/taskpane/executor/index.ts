import type {
  AddConnectorInput,
  AddLineInput,
  AddShapeInput,
  CheckFormulaConsistencyInput,
  ClearRangeInput,
  CreateArtifactInput,
  ConditionalFormatInput,
  CopyToInput,
  DataValidationInput,
  DeleteShapesInput,
  ExecuteOfficeJsInput,
  FilterRangeInput,
  FindErrorsInput,
  FindReplaceInput,
  FormatRangeInput,
  GetCellRangesInput,
  GetObjectsInput,
  InsertRowsInput,
  InsertSheetsFromFileInput,
  ListShapesInput,
  ValidateFlowchartInput,
  ManageNamesInput,
  ManageSheetInput,
  MergeRangeInput,
  ModifyChartInput,
  ModifyPivotTableInput,
  ModifyStructureInput,
  ModifyTableInput,
  PageLayoutInput,
  ReadRangeImageInput,
  ReadRangeInput,
  ResizeRangeInput,
  SearchDataInput,
  SetCellsInput,
  SetFormulaInput,
  SortRangeInput,
  ToolResult,
  UpdateShapeInput,
  WriteRangeInput,
} from "../../../protocol/src/index.js";
import { checkFormulaConsistency, findErrors } from "./audit.js";
import { findReplace } from "./replace.js";
import { copyTo, getCellRanges, readRangeImage, resizeRange, setCells } from "./cells.js";
import { toToolError } from "./common.js";
import { validateFlowchart } from "./flowchart.js";
import { filterRange, sortRange } from "./data.js";
import { formatRange } from "./format.js";
import { manageNames, mergeRange, pageLayout } from "./layout.js";
import { getObjects, modifyChart, modifyPivotTable, modifyTable } from "./objects.js";
import { executeOfficeJs } from "./office-js.js";
import { getSelection, getSheets, readRange, searchData } from "./read.js";
import { conditionalFormat, dataValidation } from "./rules.js";
import { addConnector, addLine, addShape, deleteShapes, listShapes, updateShape } from "./shapes.js";
import { insertSheetsFromFile, manageSheet } from "./sheets.js";
import { insertRows, modifyStructure } from "./structure.js";
import { clearRange, setFormula, writeRange } from "./write.js";
import { applyQaResults } from "./qa.js";
import { QaApplyResultsInputSchema } from "../../../protocol/src/index.js";

/** Tools that change the workbook; the task pane lists their results as changes. */
export const WRITE_TOOLS: ReadonlySet<string> = new Set([
  "write_range",
  "set_formula",
  "find_replace",
  "format_range",
  "clear_range",
  "insert_rows",
  "modify_structure",
  "add_shape",
  "add_connector",
  "add_line",
  "update_shape",
  "delete_shapes",
  "set_cells",
  "copy_to",
  "resize_range",
  "merge_range",
  "sort_range",
  "filter_range",
  "conditional_format",
  "data_validation",
  "manage_sheet",
  "insert_sheets_from_file",
  "manage_names",
  "page_layout",
  "modify_chart",
  "modify_pivot_table",
  "modify_table",
  "execute_office_js",
  "__qa_apply_results",
]);

/**
 * Execute one MCP tool request inside the live workbook via Office.js.
 * Arguments were already validated by the bridge. Note that Excel batches are
 * not transactional: tools validate and run their guards *before* queueing the
 * first change, so a refused call leaves the workbook untouched.
 */
export async function executeTool(tool: string, args: unknown): Promise<ToolResult> {
  try {
    return { ok: true, data: await run(tool, args) };
  } catch (error) {
    if (error instanceof UnknownTool) {
      return { ok: false, error: { code: "unknown_tool", message: error.message } };
    }
    return { ok: false, error: toToolError(error) };
  }
}

class UnknownTool extends Error {}

function run(tool: string, args: unknown): Promise<unknown> {
  switch (tool) {
    case "__qa_apply_results":
      return applyQaResults(QaApplyResultsInputSchema.parse(args));
    case "get_selection":
      return getSelection();
    case "get_sheets":
      return getSheets();
    case "read_range":
      return readRange(args as ReadRangeInput);
    case "get_cell_ranges":
      return getCellRanges(args as GetCellRangesInput);
    case "search_data":
      return searchData(args as SearchDataInput);
    case "find_errors":
      return findErrors(args as FindErrorsInput);
    case "check_formula_consistency":
      return checkFormulaConsistency(args as CheckFormulaConsistencyInput);
    case "find_replace":
      return findReplace(args as FindReplaceInput);
    case "read_range_image":
      return readRangeImage(args as ReadRangeImageInput);
    case "get_objects":
      return getObjects(args as GetObjectsInput);
    case "validate_flowchart":
      return validateFlowchart(args as ValidateFlowchartInput);
    case "list_shapes":
      return listShapes(args as ListShapesInput);
    case "write_range":
      return writeRange(args as WriteRangeInput);
    case "set_cells":
      return setCells(args as SetCellsInput);
    case "set_formula":
      return setFormula(args as SetFormulaInput);
    case "copy_to":
      return copyTo(args as CopyToInput);
    case "format_range":
      return formatRange(args as FormatRangeInput);
    case "clear_range":
      return clearRange(args as ClearRangeInput);
    case "resize_range":
      return resizeRange(args as ResizeRangeInput);
    case "merge_range":
      return mergeRange(args as MergeRangeInput);
    case "sort_range":
      return sortRange(args as SortRangeInput);
    case "filter_range":
      return filterRange(args as FilterRangeInput);
    case "conditional_format":
      return conditionalFormat(args as ConditionalFormatInput);
    case "data_validation":
      return dataValidation(args as DataValidationInput);
    case "insert_rows":
      return insertRows(args as InsertRowsInput);
    case "modify_structure":
      return modifyStructure(args as ModifyStructureInput);
    case "manage_sheet":
      return manageSheet(args as ManageSheetInput);
    case "insert_sheets_from_file":
      return insertSheetsFromFile(args as InsertSheetsFromFileInput & { fileBase64?: string });
    case "manage_names":
      return manageNames(args as ManageNamesInput);
    case "page_layout":
      return pageLayout(args as PageLayoutInput);
    case "modify_chart":
      return modifyChart(args as ModifyChartInput);
    case "modify_pivot_table":
      return modifyPivotTable(args as ModifyPivotTableInput);
    case "modify_table":
      return modifyTable(args as ModifyTableInput);
    case "add_shape":
      return addShape(args as AddShapeInput);
    case "add_connector":
      return addConnector(args as AddConnectorInput);
    case "add_line":
      return addLine(args as AddLineInput);
    case "update_shape":
      return updateShape(args as UpdateShapeInput);
    case "delete_shapes":
      return deleteShapes(args as DeleteShapesInput);
    case "create_artifact": {
      // The task pane draws the preview card from the call itself; the tool only acknowledges it.
      const { filename, content } = args as CreateArtifactInput;
      return Promise.resolve({ filename, characters: content.length, shown: true });
    }
    case "execute_office_js":
      return executeOfficeJs(args as ExecuteOfficeJsInput);
    default:
      return Promise.reject(new UnknownTool(`unknown tool '${tool}'`));
  }
}
