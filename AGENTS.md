# Repository architecture decisions

## Gateway: API for the official Claude Office add-ins

**User-confirmed requirement: the gateway is an API endpoint called by Anthropic's official Claude Office add-ins.** It is not a connection to this project's Office Partner task pane.

- The official Claude add-in runs in Excel, PowerPoint or Word and supplies its own tool schemas in API requests.
- Gateway adapts the Anthropic-compatible API to the selected provider CLI. Office tool calls become API `tool_use`; the official client executes them and returns `tool_result`.
- Preserve the client's schemas, tool names, results, conversation and continuation behavior. Support all three Office hosts; do not assume every client is Excel.
- Do not require Office Partner to be open, expose a "Gateway binding" in its Settings, ask users to select/bind a document there, or inject this project's Office tool catalog into official-client requests.
- Do not execute official-client Office tool calls through this project's Office.js executors or claim this project's backup/undo applies to actions executed by the official Claude add-in.
- Bridge serves this project's Office Partner add-in. Gateway and bridge may reuse shared provider, protocol, result-conversion and MCP management code, while retaining their different execution boundaries.
- External MCP calls hosted by gateway are distinct from official-client Office calls. They may be executed server-side using a shared MCP runtime, with explicit configuration/permission; Office tools still go back to the official client.

This decision supersedes the earlier, incorrect pane-binding implementation/design from this conversation. Do not reintroduce that design without an explicit new user requirement.

## External MCP tools: register native schemas

User-confirmed requirement: enabling an MCP server must expose its actual tools and JSON Schemas directly in the provider tool catalog. Disabling it removes those tools. Use stable server namespaces to avoid collisions. Do not expose generic list/call wrapper tools or inject MCP discovery instructions into model prompts. Refresh provider registration when the enabled catalog changes, while preserving the conversation.

## Product principles

- Neutral "Office Partner / Excel Partner" branding everywhere; never ship customer documents, company references, personal paths or account identifiers.
- Skills are optional and user-installed from `%USERPROFILE%\.excel-partner\skills`; the package bundles no skills or templates.
- Provider credentials stay under each provider CLI's control; the add-in never stores an API key. Where a provider CLI itself supports API-key sign-in (currently OpenCode), the pane may collect the key and relay it directly to that CLI's own credential store; the key must never be written to Office Partner's config, history, logs, usage records or diagnostics.
- External folders are read only when the user grants access (references / `referenceDirs`).
- Validate with `npm run typecheck`, `npm run build` and `npm test`; automated checks use mock providers, live CLI checks are opt-in.
