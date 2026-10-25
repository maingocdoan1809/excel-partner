/** Used only when the API client supplies Office.js execution tools.
 * These are model-facing preconditions, never code rewriting or local execution. */
export const OFFICE_TOOL_GUIDANCE = [
  "execute_office_js receives the BODY of an async function. context, Excel and Office are already in scope. Write statements and return a plain result directly; do not merely declare async function main(context), and do not add another Excel.run wrapper.",
  "For Office.js, explicitly load every property before reading it, including Range.address and worksheet/shape names. Queue load(...), await context.sync(), then read or serialize plain values.",
  "Inspect shape id/name/type first and sync. Only GeometricShape and TextBox are text-frame candidates; never load textFrame for all shapes, lines, pictures or groups together.",
  "Read textFrame.hasText and sync before requesting textRange.text. If an optional property fails during sync, isolate that shape and keep its identity/geometry rather than failing the entire shape inventory.",
  "Use getItemOrNullObject plus context.sync to check existence. Create a named worksheet only when isNullObject is true. Read back to verify. Do not interpret a missing/unparsed result field as a missing worksheet.",
  "A batch error may follow partial changes. Inspect current workbook state before retrying a write; do not blindly repeat the previous batch.",
  "If getSelectedRange fails with InvalidSelection (for example a shape is selected), keep workbook/sheet metadata and report that no cell selection is available. Do not pretend A1 is selected or discard the other successful reads.",
  "Excel line signature: shapes.addLine(startLeft, startTop, endLeft, endTop, Excel.ConnectorType.straight). Coordinates in points come first; connectorType is the optional fifth argument. Do not use Excel.ShapeLineType or the PowerPoint signature with a type as the first argument.",
  "Chart.getData() is not an Office.js chart API. For source inspection, use getDataRange or getDataRangeOrNullObject only if supported by the current host. Otherwise report source inspection unavailable; do not invent methods or claim independent verification from the range originally supplied to setData.",
  "For addTextBox, pass the text, then set left/top/width/height as properties. Compare chart/shape geometry with a tolerance, not exact floating-point equality.",
].join("\n");
