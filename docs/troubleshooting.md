# Troubleshooting

## Task pane shows a certificate error

The pane refuses to load over `https://localhost:8765`.

1. Confirm the cert exists: `ls certs/localhost.crt certs/localhost.key`.
2. Re-import `certs/localhost.crt` into **Trusted Root Certification
   Authorities** for the current user.
3. Reopen Excel so the WebView2 cache is cleared.

## Task pane stays on `connecting` or `closed`

- Verify the bridge: `curl -k https://localhost:8765/health`.
- Check the log: `run/bridge.log`.
- Windows must reach WSL `localhost`. If it does not, set
  `BRIDGE_HOST=0.0.0.0` before starting and use the WSL IP instead.
- The bridge only answers requests whose `Host` is `localhost`, `127.0.0.1`,
  `::1` or its configured host, and refuses other web origins with `403`
  (`rejected request from untrusted origin or host` in the log). To reach it
  by another name such as a WSL IP, allow that name and restart:
  - npm install: `excel-partner allow-host 172.20.0.5`, then `excel-partner restart`
    (`excel-partner allow-host` alone lists them; `--remove` removes one)
  - source install: re-run the installer with `-AllowedHost 172.20.0.5`
    (`scripts/install.ps1`) or `--allowed-host 172.20.0.5`
    (`scripts/install-windows.sh`), or add `"allowedHosts": ["172.20.0.5"]` to
    `excel-partner.config.json`
  - one-off: `BRIDGE_ALLOWED_HOSTS=172.20.0.5,other-name` (comma separated)

## Provider shows `not_installed` / `login_required`

- Install the CLI **inside WSL**, not on Windows.
- Log in once in a WSL terminal (`claude`, `copilot`, `codex login`, `opencode auth login`).
- Point the bridge at a non-default binary with `CLAUDE_BIN`, `COPILOT_BIN`,
  `CODEX_BIN` or `OPENCODE_BIN`.

### Provider-specific notes (verified versions)

| Provider | Tested with | How the bridge drives it | Typical failure |
| --- | --- | --- | --- |
| Claude Code | 2.1.294 | one `claude -p --output-format stream-json` process per turn, `--resume` between turns | `login_required` + "organization has disabled ..." means the API rejected the account, not the add-in; use another provider or an `ANTHROPIC_API_KEY` |
| Copilot CLI | 1.0.95 | `copilot --acp` over stdio, Excel MCP server via `--additional-mcp-config @file` | `not_installed` / `login_required`: run `copilot login` |
| Codex CLI | 0.162.1 | `codex app-server` over stdio, Excel MCP server via `thread/start` config | `login_required`: run `codex login` |
| OpenCode CLI | 1.19.3 | `opencode acp` over stdio, Excel MCP server via `session/new` `mcpServers` | `login_required`: run `opencode auth login`; free OpenCode models work without any login |

A version older than the tested minimum is reported as `error` with an update hint.
Only the seven Excel tools are exposed to the model; shell, file and web tools of
every CLI are disabled or denied. Set `BRIDGE_LOG_LEVEL=debug` to see the raw CLI
stderr and Copilot ACP updates in the bridge log.

## The AI says it cannot draw shapes / lacks a tool

The task pane or the provider session is older than the bridge. After an update: restart the bridge
(`bash scripts/start.sh` after `stop.sh`), then reload the pane (close it and click **Home > Excel Partner**,
or right-click the pane > Reload) and switch provider or press **Clear** so a new provider session
picks up the current tool list. Shapes need ExcelApi 1.9; cell notes need ExcelApi 1.18.

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

After updating, restart the bridge, reload the add-in, and start a new provider session so the new tool is exposed.

## The assistant says it cannot open a file by path

File access is opt-in: add the folder to `referenceDirs` in `excel-partner.config.json` and restart the bridge
(`bash scripts/stop.sh; bash scripts/start.sh`). Only files inside that folder, read-only.

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

Excel did not answer within `BRIDGE_TOOL_TIMEOUT_MS` (default 30s). Large
ranges, a modal dialog in Excel, or a busy UI can cause this. Reduce the
selection size or retry.

## `size_limit_exceeded`

The request exceeded the read/write cell limits. Tune with
`BRIDGE_MAX_READ_CELLS` / `BRIDGE_MAX_WRITE_CELLS` in the bridge environment.

## Reset local state

```bash
bash scripts/stop.sh
rm -rf run/ && bash scripts/start.sh
```
