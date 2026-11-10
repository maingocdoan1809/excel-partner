# Troubleshooting

Install, certificate, connection and update problems are covered in the README ("Troubleshooting" section). This page explains behavior you may see while the assistant works. Restarting the bridge means `office-partner restart`; from a source checkout it means `bash scripts/stop.sh && bash scripts/start.sh`.

## The AI says it cannot draw shapes / lacks a tool

The task pane or the provider session is older than the bridge. After an update: restart the bridge, reload the pane (close it and click **Home > Office Partner**, or right-click the pane > Reload) and start a new chat so the provider session picks up the current tool list. Shapes need ExcelApi 1.9; cell notes need ExcelApi 1.18.

## `get_selection` says a shape or chart is selected

On hosts supporting ExcelApi 1.19, `get_selection` returns `activeShape` with its ID, name, text, geometry and connector endpoints. This identifies one active shape. A Windows bridge additionally reads the selected group through COM: `selectedShapesStatus: "complete"` means `selectedShapes` contains the selected names, native IDs, native types and connector flags. The same reader enriches the message context sent to the AI. Native IDs may differ from Office.js IDs; use shape names with tools and `list_shapes` for details. This read is independent of `execute_office_js`.

COM reading attaches to Excel's actual workbook window using `AccessibleObjectFromWindow` (`EXCEL7` / `OBJID_NATIVEOM`), avoiding stale `GetActiveObject` registrations. It requires exactly one matching workbook window, checked by workbook name, worksheet and the workbook path when provided. A mismatch, ambiguous windows, busy Excel or timeout falls back to Office.js and reports `selectedShapesStatus: "unavailable"`; the model must not treat one active shape as the full group. Other bridge platforms retain the Office.js behavior. Older Office.js hosts return `shapeSelectionStatus: "unavailable"` when identifying an object requires ExcelApi 1.19. See [Microsoft's active-shape API](https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-1.19#excel-excel-workbook-getactiveshapeornullobject-member(1)).

The selection label above the prompt uses native `Shape.onActivated` / `onDeactivated` events (ExcelApi 1.9), alongside document selection events for cells. There is no periodic selection polling. Shape event subscriptions refresh on worksheet activation, successful shape tools, and pane focus. Group children are included. Selection reads fetch only identity/range dimensions and do not change the prompt text.

`list_shapes` includes nested group children, parent IDs, text, geometry, and actual connector attachments and arrows. Follow `nextOffset` while `hasMore` is true; the default page has 200 shapes and `limit` can be increased to 1,000.

## Validate flowchart connections and connector overlaps

Ask the assistant: **Validate this flowchart: find disconnected shapes, crossing or overlapping connectors, and lines going through other shapes. Report the affected shape names before making changes.**

The read-only `validate_flowchart` tool scans up to 1,000 shapes on the active sheet (or the supplied `sheet`). It reports structural clues and geometric issues, including crossing straight connectors, shared line segments, overlapping node bounds, and connectors passing through unrelated node bounds. Shared connector endpoints are allowed. Rectangle bounds are a screening model; they do not describe every curved shape outline.

Elbow/curve connector routes, unsupported connection sites, grouped coordinates, and rotations require visual review; the tool reports incomplete coverage rather than inventing routes. `no_issues_in_checked_geometry` only describes the checked model, not business correctness. `scanTruncated` means the entire sheet was not checked. The tool does not move shapes or reroute connectors. After requested layout changes, run it again.

If Excel rejects a text read (for example, `TextRange._OnAccess`), the tools keep the shape geometry and connections. They check `hasText` before opening a text range and retry rejected text batches individually. The affected shape has `textReadStatus: "unavailable"`, `text: null`, and a `readWarnings` entry naming the failed property. Empty supported shapes return `text: ""` and `textReadStatus: "empty"`; lines have `not_applicable`. Both tools expose `textReadComplete` separately from geometric validation. A text warning does not justify deleting a shape or claiming that the whole sheet cannot be checked.

## The assistant says it cannot open a file by path

File access is opt-in: add the folder as a reference from the composer's **@** menu, or add it to `referenceDirs` in the configuration file and restart the bridge. Only files inside that folder, read-only.

## A write is refused with `would_overwrite`

By design. The tool names the non-empty cells; the AI should ask you, then retry with `allowOverwrite`.
Say "replace" / "overwrite" in your request to skip the question.

## A turn does not stop after pressing Stop

The bridge asks the provider to cancel, and after `BRIDGE_CANCEL_GRACE_MS`
(default 8s) it reports the turn as cancelled and drops the provider process so
the next prompt starts clean.

## A tool call fails with `disconnected`

The task pane closed or reloaded mid-turn. The bridge cancels pending tool calls
and does **not** replay ambiguous writes. Re-send the request once the pane is
reconnected.

## A tool call fails with `timeout`

Office did not answer within `BRIDGE_TOOL_TIMEOUT_MS` (default 120s). Large
ranges, a modal dialog in the Office app, or a busy UI can cause this. Reduce the
selection size or retry.

## `size_limit_exceeded`

The request exceeded the read/write cell limits. Tune with
`BRIDGE_MAX_READ_CELLS` / `BRIDGE_MAX_WRITE_CELLS` in the bridge environment.
