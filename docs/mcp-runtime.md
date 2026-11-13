# Shared MCP core and gateway API

Bridge and gateway share MCP implementation code, while Office execution stays with the add-in that supplied the tools:

```
Task pane / provider MCP -> OfficeRuntime -> Office Partner pane -> host executor
Official Claude Office add-in -> Gateway API -> provider CLI
Official Claude Office add-in <- tool_use / tool_result -> Gateway API

Shared MCP connection management: McpServers -> stdio / Streamable HTTP
```

## Ownership

- `protocol/src`: tool input schemas, host availability, messages and result contracts.
- `runtime/src/catalog.ts` and `tool-definitions*.ts`: the shared advertised catalog. `excel_mcp/src` is only the stdio MCP entry point and its HTTP forwarder.
- `runtime/src/office-runtime.ts`: validation, host/capability gates, dispatch and external-tool approval. Existing domain helpers/session models remain under `bridge/src`; they are reused, not copied into the gateway.
- `runtime/src/mcp-servers.ts`: persisted server configuration and per-pane MCP connection lifecycle.
- `runtime/src/results.ts`: shared conversion for external MCP results; gateway retains its separate API wire conversion for official-client tool results.
- The Office Partner task pane owns its host executor, Ask/Auto prompts. Official Claude clients own their independent Office executors.

No provider gets an external server’s command configuration. Enabled servers expose their actual tools and input schemas directly, using stable names such as `ext_<server namespace>__summarize_numbers`. Disabled servers contribute no tools. There are no generic list/call wrapper tools or discovery prompts.

## MCP setup

1. Open Settings, then **MCP servers → Add MCP server**.
2. Enter a local command with a JSON array of arguments, or a Streamable HTTP URL. Save does not launch the server.
3. **Test connection** explicitly launches/connects and discovers tools. **Edit** updates a server; editing removes its enabled selection from existing panes so it must be enabled again.
4. In the composer, open **Tools → Tools** and enable the server for this connected file. Selections are remembered in sessionStorage for this pane tab and restored after a reconnect/reload. New tabs are not enabled automatically, and changed server configurations require enabling again.
5. In Ask mode, each external tool call shows the server, tool and arguments. Schema validation uses the tool schema returned by the selected server. Server output is data, not instructions.

Configuration is stored in `mcp-servers.json` beside the bridge's config file. Server IDs are distinct from display names. Connections are isolated by pane and closed when the pane disconnects. Updating/removing a server closes its connections. The bridge limits discovery to 500 tools / 1 MB and returned call results to 4 MB. Calls time out after 60 seconds and are never replayed automatically.

Local MCP processes run as the current user; the bridge passes a small environment allowlist rather than provider/bridge credentials. MCP metadata such as read-only hints is not treated as permission. External MCP side effects are outside Office's backup and undo.

This first version supports stdio and unauthenticated Streamable HTTP. It does not implement OAuth, custom auth headers/environment secrets, legacy SSE, resources/prompts, or MCP server-initiated sampling/elicitation. Servers needing these features will not work merely by entering a URL. Use trusted servers; local command arguments are stored in the local configuration and should not contain secrets.

## Gateway is an API for the official Claude Office add-ins

This is a user-confirmed architecture requirement; see `AGENTS.md`.

The gateway is an Anthropic-compatible API endpoint used by Anthropic’s official Claude add-ins for Excel, PowerPoint and Word. The official client supplies tool schemas, executes Office actions and returns tool results. Gateway adapts API requests to a provider CLI, returns `tool_use` to the official client, then resumes the provider when `tool_result` arrives.

Gateway does not require the Office Partner pane, does not bind to its documents and must not expose a Gateway binding control in Office Partner Settings. It does not inject Office Partner’s tool catalog into official-client requests. Office Partner’s backup/undo cannot be claimed for actions performed by the official add-in.

Bridge and gateway can reuse MCP connection management, schemas, provider transports and result conversion. External MCP tools hosted by gateway are a separate execution category: the shared MCP runtime can execute those server-side, while official-client Office tools always return to the official client.

The former pane-binding design was based on a misunderstanding and has been removed.

## Gateway-managed external MCP configuration

For external MCP servers, gateway reuses `McpServers` and `invokeMcpTool` from `runtime/src`, with a gateway configuration policy instead of Office Partner UI approvals. Example gateway `config.json`:

```json
{
  "mcpConfigFile": "C:\\Users\\you\\AppData\\Local\\Excel Partner\\mcp-servers.json",
  "enabledMcpServers": ["your-server-id"],
  "mcpApprovalMode": "auto"
}
```

Use the actual path of your server configuration. The same file may be managed through Office Partner Settings, but a running Office Partner bridge or pane is not required by gateway. An independent configuration file with the same format works as well. Server IDs are visible in the configuration file. Gateway reloads it before API requests and closes connections when the configuration changes.

`mcpConfigFile` defaults to `null`, `enabledMcpServers` to an empty list, and `mcpApprovalMode` to `deny`. Set `auto` only for servers the gateway administrator authorizes; there is no Office Partner approval popup in the official Claude add-in. Office actions remain subject to the official client's own behavior. External tool namespaces are registered only for enabled servers; collisions with client-supplied tool names are rejected.

Gateway registers the actual tools and input schemas from enabled servers, under collision-safe namespaces. It executes those external calls itself. All other, official-client tools continue through `tool_use` and `tool_result`. It never copies Office Partner tool schemas into the official client's Office catalog.

Regression coverage: `tests/gateway-mcp.test.ts` runs this complete cycle for Excel, Word and PowerPoint without starting a bridge or Office Partner pane, including server-side MCP results, client-owned Office results, policy denial and name collisions.

## Provider catalog refresh and verification

Claude is launched with `--system-prompt-snapshot off`: it rebuilds the current prompt/tool definitions rather than restoring a snapshot that retains disabled tools. Native conversation messages and session identity are preserved; no MCP discovery instructions are added.

Copilot disables persisted MCP tool snapshots (`disableToolCache: true`) and tool deferral (`deferTools: "never"`, `toolSearch.enabled: false`). Its existing available-tools filter still limits callable tools. This ensures the live MCP catalog is presented directly after resume instead of depending on stale or deferred tool discovery. Models that reject reasoning effort are retried at session creation without that unsupported setting, keeping the requested model and native session identity.

Run an opt-in model-backed check with `npx tsx scripts/live-mcp-smoke.ts <provider> <model>` after building. It uses an isolated temporary configuration and real stdio fixture, records completed server calls, checks on/off/on selection and native conversation identity, and removes its temporary files afterward. It can incur provider usage; it is not part of the automatic suite. The requested model is passed explicitly; the provider service may apply its own routing or fallback.

Verified with Claude Haiku, OpenCode `opencode/big-pickle`, and Copilot (requested `gpt-4.1`): real external calls succeed before and after toggling off, no external call executes while disabled, and the native conversation remains the same. The strict real-Claude/fake-API test also inspects outgoing request schemas and preserved history. `tests/copilot-mcp-catalog.live.test.ts` inspects the real Copilot runtime's schemas through on/off/on using fresh runtimes without making a model request. Resume identity and history are covered by the opt-in model-backed smoke; Copilot does not persist a resumable conversation before its first model turn.
