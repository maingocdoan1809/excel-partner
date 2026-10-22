# Windows COM selection POC

`poc-shape-selection.ps1` reads the current Excel window's selected shape IDs,
names, types and connector flags. It emits JSON on changes. It does not edit
the user's workbook and is not connected to the add-in UI or AI prompt.

Run from Windows PowerShell in this repository:

```powershell
& ([scriptblock]::Create([IO.File]::ReadAllText("$PWD\scripts\poc-shape-selection.ps1"))) -WatchSeconds 90
```

While watching, select a shape, Ctrl+click additional shapes/connectors,
remove some from the selection, and return to a cell. Compare `count` and
`shapes` with the Excel selection. With no arguments it reads once.

For a repeatable test in a separate, hidden Excel application:

```powershell
& ([scriptblock]::Create([IO.File]::ReadAllText("$PWD\scripts\poc-shape-selection.ps1"))) -SelfTest
```

Self-test creates an unsaved temporary workbook with two shapes and a connector,
checks six selections, then closes that workbook without saving and quits only
the Excel instance it created. All six cases passed on the development machine:
single shape; two shapes; two shapes plus connector; one shape plus connector;
single connector; cell C5.

The read-only probe also connected to the user's running workbook and identified
its worksheet and cell selection. Interactive multi-selection changes in that
workbook have not yet been verified.

Before production integration:

- Resolve the specific workbook/window belonging to the add-in. `GetActiveObject`
  alone does not reliably choose between multiple Excel processes.
- Handle Excel being busy, closed, editing a cell or switching worksheets.
- Verify manual Ctrl+click, selection-box and group-child behavior.
- Measure latency and resource use of repeated reads before choosing a polling
  interval. The POC requests 500 ms between reads, not a guaranteed response time.
- Keep Office.js as the fallback on platforms without Windows COM.

Reference: [Excel ShapeRange](https://learn.microsoft.com/en-us/office/vba/api/excel.shaperange).
