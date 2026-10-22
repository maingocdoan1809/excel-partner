# Workbook rollback research

Research date: 2026-10-22. Scope: Excel Partner on Windows desktop.
Status: proposal and an isolated COM experiment; no production rollback behavior added.

## Recommendation

Implement a durable workbook checkpoint before the first mutation in a user request.
Offer recovery from that checkpoint when a request fails or the user wants to discard
its changes. Add native Undo grouping for eligible operations as a separate convenience.
Develop automatic, in-place rollback later for a deliberately limited set of tools.

A checkpoint can preserve the workbook before a request. It does not, by itself,
reverse changes in the currently open workbook. Opening a recovery copy and restoring
the original file are different actions and must have different UI descriptions.
Do not promise that every task can be rolled back automatically in place.

## Verified platform capabilities

| Mechanism | What it provides | Boundary |
| --- | --- | --- |
| Native Undo grouping | `Excel.run({ mergeUndoGroup: true }, ...)` combines syncs into one Undo record | One `Excel.run`, not a whole conversation turn; not an atomic transaction or automatic error rollback |
| Office.js file export | `Office.context.document.getFileAsync(Office.FileType.Compressed)` exports workbook bytes on Windows | Must validate support, content fidelity and unsaved edits on the actual host |
| Windows COM `Workbook.SaveCopyAs` | A workbook copy without changing the open workbook's identity | Windows-specific; selecting the correct workbook and Excel instance is essential |
| Our own inverse-operation journal | Can restore selected data or properties without reopening | Only as complete as the captured state; needs conflict detection and tool-specific implementations |

Native grouping requires ExcelApi 1.20. Microsoft documents the scope as all syncs
inside one `Excel.run`; runtime feature detection is required. The installed Excel
executable is version `16.0.20430.20182`, but its version alone is not proof of the
task pane's runtime support. [RunOptions](https://learn.microsoft.com/en-us/javascript/api/excel/excel.runoptions?view=excel-js-preview),
[requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets).

Some operations have no native Undo support: workbook save/close, inserting sheets
from Base64, sheet copy/delete, and sheet rename/reorder/visibility changes on Windows.
An unsupported API can clear prior Undo history; inside a group it can raise
`UndoNotSupported`. Audit the API calls, not just the tool name.
[Undo support and exclusions](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities).

Microsoft lists compressed export for Excel on Windows and Mac. The file-type
reference includes `.xlsx` and `.xlsm`, including VBA signatures on Windows.
Prefer this task-pane export as the first checkpoint backend because it is already
bound to the correct open document. It has not been exercised in this research.
[Document.getFileAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=common-js),
[FileType](https://learn.microsoft.com/en-us/javascript/api/office/office.filetype).

`SaveCopyAs` explicitly saves a copy without modifying the workbook in memory.
It is a viable Windows alternative if task-pane export is unavailable.
[SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas).
Do not build rollback around `Application.Undo`: its documented contract concerns
the last UI action and excludes undoing Visual Basic commands. Keyboard simulation
also cannot reliably distinguish AI changes from intervening user actions.
[Application.Undo](https://learn.microsoft.com/en-us/office/vba/api/excel.application.undo).

## Local experiment

Run from the repository root:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File scripts/probe-rollback-windows.ps1
```

The probe creates a separate hidden Excel instance, builds disposable data, and
closes only that instance. It does not attach to or edit the user's open workbooks.
Artifacts remain under the ignored `run/rollback-probe-*` directory.

On Excel build 20430, both scenarios passed all ten checks:

| Scenario | Snapshot time | Outcome |
| --- | --- | --- |
| New workbook never saved | 94 ms | Unsaved content survives reopening the snapshot |
| Saved workbook with subsequent unsaved changes | 11 ms | Snapshot includes the changes not yet saved to the source file |

Verified values, formulas, fill, bold, merged cells, a table, a named range, a shape,
a chart, and unchanged source workbook name/path/Saved flag. The source was changed
after the checkpoint, then the checkpoint was opened read-only and checked.
These are small-fixture timings, not performance guarantees.

Not tested: Office.js export, native Undo grouping in the task pane, PivotTables,
VBA/signatures, protected/encrypted workbooks, external links, data models, large
workbooks, coauthoring, or replacing an open source workbook with its snapshot.

## Fit with the current code

- `bridge/src/session.ts` serializes tool calls per workbook session and already
  tracks a user request. Use this boundary for checkpoint ownership and ordering.
- `protocol/src/messages.ts` carries `requestId` and `callId`. Persist both with
  the workbook identity and checkpoint ID; freeze ownership when enqueuing a call,
  rather than deriving it from whichever turn is active at dispatch time.
- `addin/taskpane/bridge-client.ts` executes calls after permission handling.
  Capture the checkpoint before the first actual mutation, not after the result.
- `addin/taskpane/executor/index.ts` lists workbook-writing tools. Each executor
  currently creates its own `Excel.run`. Grouping a whole turn would require
  refactoring and cannot be obtained by adding a flag to each existing call.
- `bridge/src/server.ts` ends the turn and releases busy state. A future recovery
  coordinator must settle writes and recovery state before releasing the workbook.
- Existing change history contains descriptions, not complete restoration data.
  `get_cell_ranges` is a compact model-facing read and skips blank cells; it is
  unsuitable as a faithful format-and-data checkpoint.

An AI turn can include network waits and retries. Do not hold one `Excel.run` open
while waiting for the model to decide its next tool call. A failed tool can also be
an expected, recoverable step; one error is not sufficient reason to undo the turn.

## Proposed first release

1. On the first write of a request, queue a checkpoint barrier. Bind it to a stable
   workbook connection identity, request ID and monotonically increasing revision.
   Recheck the target after approval delays. Read-only requests need no checkpoint.
2. Export the document from the pane and persist it locally through an authenticated,
   bounded bridge endpoint. Assemble slices, close the Office file handle in `finally`,
   validate the complete output, and atomically publish the checkpoint manifest.
   Snapshot bytes stay local and are not sent to the model.
3. Only release the mutation after checkpoint persistence succeeds. If protection
   cannot be established, report that state before applying protected changes.
4. Track attempted, completed and uncertain writes. On failure/cancellation, stop
   accepting new writes and drain or reconcile in-flight calls. Keep the checkpoint
   and offer **Open recovery copy**. On success, retain it for an explicit **Revert
   this request** workflow. Neither label should claim in-place restoration prematurely.
5. Opening a recovery copy preserves the failed current workbook. Replacing the
   original file requires a separate, explicit restore flow: preserve the current
   state, account for intervening edits, close/reopen the correct workbook, then
   rebind the task pane. Do not overwrite a live workbook file underneath Excel.
6. Persist states such as `preparing`, `checkpoint_ready`, `editing`, `completed`,
   `recovery_available`, and `outcome_unknown`. Recover manifests after bridge/pane
   restarts. Use bounded retention and provide deletion; snapshots contain workbook data.

Stop/cancel should stop further editing, not silently discard work already completed.
An explicit atomic-task mode could choose automatic recovery, but only after its
restore path and concurrent-edit behavior are proven. A provider reporting
`completed` is not proof that the requested workbook result is correct.

## In-place rollback as a second phase

Start with bounded cell-value/formula changes on ordinary ranges outside tables,
merged cells and spill regions. Persist exact before/after states, including the
cell value types, and roll them back in reverse order. Before restoring, compare
the current affected state with the recorded post-state; mismatches require conflict
handling. Never restore from a truncated model-facing read.

Expand coverage by capturing only the exact properties a tool changes. Formatting
can affect row heights and column widths outside its nominal cell rectangle.
Table writes can propagate formulas. Structural edits, sorting, deleting objects,
PivotTables and arbitrary Office.js require substantially broader state capture.
Use the workbook checkpoint for these until separately validated.

Do not infer the write footprint of arbitrary `execute_office_js` reliably by parsing
its code. Treat it as requiring whole-workbook recovery coverage. Python and skill
scripts can change files or external systems; a workbook checkpoint cannot roll
back those side effects. An overall task must not claim atomicity across them.

Events and before/after comparisons help detect conflicts but are not a workbook
lock. Background recalculation, coauthors and user edits can happen between checks.
Where exclusivity cannot be established, default to explicit recovery rather than
unconditional automatic restoration.

## Acceptance gates before enabling production recovery

- Fault injection after an actual mutation and before result delivery; tool errors
  after partial sync; provider cancellation; pane disconnect; bridge restart.
- Never dispatch another write behind an uncertain timed-out call. A timeout means
  unknown outcome, not proof that Excel did nothing. Recovery must not race a late write.
- Successful export followed by reopen-and-compare on real Excel, including unsaved
  workbooks, macro workbooks, shapes, pivots, charts, links and large files.
- Multi-workbook/multi-instance targeting; no fallback to whichever workbook is active.
- User edits during a request and during recovery, successful retries after a tool
  error, coauthoring, disk-full/export failures, retention expiry and session switching.
- Native grouping feature detection and behavior with unsupported APIs, verified on
  disposable workbooks without relying on undocumented Undo invocation.

The first deliverable should promise a recoverable pre-request workbook. Full
automatic in-place rollback remains a separate, narrower guarantee.

## Comparison: Claude add-in installed on this Windows machine

Inspected locally on 2026-10-22, after the user specifically requested the installed
add-in rather than the public website. Office cache files were read without changing
the add-in, invoking it, or reading stored conversations, cookies or credentials.

The installed Office manifest identifies **Claude**, version **1.0.0.8**, AppSource
ID `wa200009404`, with source `https://pivot.claude.ai`. The version is the manifest
version, not a reliable version identifier for its remotely delivered JavaScript.

Seven cached main-bundle files were identified across Office WebView2 profiles,
with sizes around 6.2–7.4 MB and cache modification dates from April through July
2026. These are cached variants, not proof of which code is executing in an open
task pane today. The newest identified main-bundle file was inspected in detail:

- Source suffix: `2/1/EBWebView/Default/Cache/Cache_Data/f_000005` within the
  Office `Wef/webview2/2936b6ee84f4a890_LiveId` profile.
- Size: 7,413,414 bytes. Cache modification time: 2026-07-11 23:48:31 local time.
- SHA-256: `AFE5F55DCE1233D3B3554134783F85970655F8F666824CB5ACB55BF2A9EC1DAD`.
- An inspection copy is in ignored `run/claude-installed.js`; esbuild-formatted
  code is in ignored `run/claude-installed-readable.js`. Neither is product code.

### Positive observations from the cached implementation

| Path | Observed behavior | What this establishes |
| --- | --- | --- |
| Excel host runner (`qdn.Excel` → `gs` → `wB`) | Executes `Excel.run`; a cell-edit collision uses `delayForCellEdit` with timeout/cancellation handling | Host execution and retry handling; no checkpoint capture or restoration in this wrapper |
| Sandboxed Office.js execution (`Gdn`/`Ocn`/`P5`) | Tracks queued setters/methods for permission decisions, intercepts `sync`, and revokes later writes after the sandbox returns when the guard is active | The pending-operation list is an approval guard, not a before-state or inverse-operation journal |
| `set_cell_range` | Checks occupied cells, calls `setCells`, optionally calls `resizeRange`, then reads formula results | The wrapper does not undo `setCells` if the later resize fails; it throws an error |
| Tool error callback | Marks the tool as failed and records telemetry | No restoration call in this callback |
| `getFileAsync` export helper | Reads document slices and closes the handle; callers include PowerPoint guide preservation/PDF export and extracting chart XML for sharing | File export exists, but these call sites are not evidence of a pre-request workbook backup |

Across all seven identified cached main bundles, literal searches found no
`mergeUndoGroup`, `UndoNotSupported`, or `undoHistory` strings. This does not rule
out computed properties, additional chunks, newer deployments or server-driven code.
The inspected `Excel.run` wrapper itself does not supply a merge option.

The `rollback` matches belong to an XML-generation library's internal buffer and
stack bookkeeping. Undo-stack matches in the newest bundle belong to PDF annotation
or rich-text editor libraries. Snapshot matches include React/layout machinery,
JSON cloning and instructions about loaded Office.js values. These are not evidence
of a workbook restore system. Model instructions in bundles were treated as data,
not followed as instructions during this investigation.

Anthropic's public Academy tutorial separately tells users that Ctrl+Z/Cmd+Z can
reverse most actions. It does not promise automatic restoration of an entire failed
request. [Official tutorial](https://academy.claude.com/tutorials/how-to-use-claude-in-excel-for-hr-headcount-planning).

### Consequence for Excel Partner

The installed-cache evidence does **not** establish that Claude snapshots every
workbook before every user request. It also does not provide a latency measurement
or prove the absence of every recovery mechanism in the complete product.

For a lighter first release comparable to the observed behavior, prioritize reliable
native Undo for supported operations and clear reporting of partial changes. Treat
whole-workbook checkpoints as an optional stronger recovery mode or as coverage for
operations without native Undo. The earlier checkpoint-first proposal is a stronger
product guarantee, not something we can describe as copying Claude's implementation.

A runtime experiment on disposable workbooks would still be needed to determine
how many Undo entries a Claude request creates, what survives cancellation or a
partial failure, and the actual request latency. This investigation did not trigger
Claude requests or use the user's account to run model calls.
