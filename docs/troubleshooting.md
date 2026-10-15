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

## Provider shows `not_installed` / `login_required`

- Install the CLI **inside WSL**, not on Windows.
- Log in once in a WSL terminal (`claude`, `copilot`, `codex login`).
- Point the bridge at a non-default binary with `CLAUDE_BIN`, `COPILOT_BIN` or
  `CODEX_BIN`.

### Provider-specific notes (verified versions)

| Provider | Tested with | How the bridge drives it | Typical failure |
| --- | --- | --- | --- |
| Claude Code | 2.1.294 | one `claude -p --output-format stream-json` process per turn, `--resume` between turns | `login_required` + "organization has disabled ..." means the API rejected the account, not the add-in; use another provider or an `ANTHROPIC_API_KEY` |
| Copilot CLI | 1.0.95 | `copilot --acp` over stdio, Excel MCP server via `--additional-mcp-config @file` | `not_installed` / `login_required`: run `copilot login` |
| Codex CLI | 0.162.1 | `codex app-server` over stdio, Excel MCP server via `thread/start` config | `login_required`: run `codex login` |

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

When a shape/chart is selected Excel cannot return a cell range. The tool reports `note: "A shape, chart or other
object is selected"` (no error); select the cells you mean, or ask the assistant to use `list_shapes` / `get_objects`.

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
