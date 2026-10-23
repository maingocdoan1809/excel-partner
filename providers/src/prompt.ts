import type { SelectionSnapshot } from "../../protocol/src/index.js";

/** No add-in instructions by default; providers retain their own native instructions. */
export const SYSTEM_PROMPT = "";

const MAX_PROMPT_VALUES = 400;

/** Combine the user's request with the selection snapshot taken when they hit send. */
export function buildTurnPrompt(prompt: string, selection: SelectionSnapshot | null): string {
  if (!selection) return prompt;

  const values = selection.values.flat().length > MAX_PROMPT_VALUES
    ? selection.values.slice(0, Math.max(1, Math.floor(MAX_PROMPT_VALUES / Math.max(1, selection.columnCount))))
    : selection.values;

  const context = {
    workbook: selection.workbookName,
    sheet: selection.sheet,
    address: selection.address,
    rows: selection.rowCount,
    columns: selection.columnCount,
    headers: selection.headers,
    values,
    truncated: selection.truncated || values.length < selection.values.length,
    ...(selection.activeShape ? { activeShape: selection.activeShape } : {}),
    ...(selection.selectedShapes ? { selectedShapes: selection.selectedShapes } : {}),
    ...(selection.selectedShapesSource ? { selectedShapesSource: selection.selectedShapesSource } : {}),
    ...(selection.selectedShapesStatus ? { selectedShapesStatus: selection.selectedShapesStatus } : {}),
    ...(selection.selectedShapesNote ? { selectedShapesNote: selection.selectedShapesNote } : {}),
    ...(selection.shapeSelectionStatus ? { shapeSelectionStatus: selection.shapeSelectionStatus } : {}),
    ...(selection.note ? { note: selection.note } : {}),
  };
  return [
    "<selection_snapshot>",
    JSON.stringify(context),
    "</selection_snapshot>",
    "(The snapshot was taken when the user sent this message; call get_selection for the live state.)",
    "",
    prompt,
  ].join("\n");
}
