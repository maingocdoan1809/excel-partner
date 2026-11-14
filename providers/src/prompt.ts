import type { PromptSelectionSnapshot, SelectionSnapshot } from "../../protocol/src/index.js";

/** No add-in instructions by default; providers retain their own native instructions. */
export const SYSTEM_PROMPT = "";

const MAX_PROMPT_VALUES = 400;

/** Wrap a snapshot object and the user's request the way every host does. */
function withSnapshot(context: object, prompt: string, note = ""): string {
  return [
    "<selection_snapshot>",
    // The snapshot carries text from the file (cells, document prose, slide text). "<" is written as \u003c so that
    // text like "</selection_snapshot>" inside a file cannot close the block and pass for the user's own words.
    JSON.stringify(context).replace(/</g, "\\u003c"),
    "</selection_snapshot>",
    `(The snapshot was taken when the user sent this message; call get_selection for the live state.${note})`,
    "",
    prompt,
  ].join("\n");
}

function excelSnapshotContext(selection: SelectionSnapshot): object {
  const values = selection.values.flat().length > MAX_PROMPT_VALUES
    ? selection.values.slice(0, Math.max(1, Math.floor(MAX_PROMPT_VALUES / Math.max(1, selection.columnCount))))
    : selection.values;

  return {
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
}

/** The snapshot names the one active shape; the whole selected group (and each shape's identity) is read when the model asks for it. */
function excelSnapshotNote(selection: SelectionSnapshot): string {
  return selection.activeShape && !selection.selectedShapes
    ? " A shape is selected: call get_selection to see every selected shape." : "";
}

/**
 * Combine the user's request with the selection snapshot taken when they hit send. A snapshot without a `host` is
 * Excel's; Word and PowerPoint snapshots carry no cell values, so they are passed on as the pane sent them.
 */
export function buildTurnPrompt(prompt: string, selection: PromptSelectionSnapshot | null): string {
  if (!selection) return prompt;
  if ("workbookName" in selection) return withSnapshot(excelSnapshotContext(selection), prompt, excelSnapshotNote(selection));
  return withSnapshot(selection, prompt);
}
