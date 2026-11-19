# AGENTS.md

Guidance for coding agents (Claude Code, Codex CLI, Copilot CLI, …) working in this repository. `CLAUDE.md` imports this file; edit this one.

## What this is

Office Partner (npm `office-partner`, formerly Excel Partner): a Windows desktop add-in for Excel, Word and PowerPoint that chats with a local provider CLI (Claude Code, Copilot CLI, Codex CLI, OpenCode). The task pane (Office.js) talks to a local HTTPS/WebSocket **bridge** on `localhost:8765`; the bridge drives the provider CLI, which calls Office tools through a stdio MCP server that forwards back to the bridge and on to the pane. A separate **gateway** (`localhost:8766`) is an Anthropic Messages-compatible API backed by the same provider CLIs.

## Working rules (from the owner)

- **Interfaces first, no duplication.** Build on the existing contracts (`ProviderAdapter`, `BaseCliAdapter`, `LoginAttempt`, the shared tool definitions, `McpServers`, …) and extend them instead of adding parallel helpers. Before writing new code, look for an existing function or base-class method that already does it.
- **Everything provider-specific goes through the provider.** Never branch on a provider name (`if (provider === "copilot")`) outside `providers/`. Add a method or a `capabilities` flag to the interface, implement it in each adapter (default in `BaseCliAdapter`, override where a provider differs), and have callers use only the interface.
- **Self-review when done.** After coding, review your own diff against these rules (interface reuse, no duplication, no provider branches, layer direction) and report the verdict, along with typecheck/test results.
- **Try changes locally with `npm run dev`.** Build only when it is actually needed. To install a build on this machine, use the add-in's own update mechanism: `npm run pack:release` (→ `release/office-partner-<version>.tgz`), then run the updater from the repo, `node bin/office-partner.mjs update --from release/office-partner-<version>.tgz`. Don't use the installed `office-partner update` for this, since it runs the old code.
- Reply to the owner in Vietnamese, concisely.

## Commands

Node must be `^20.19 || >=22.12 <26`. On the owner's machine Node 26 is first on PATH; run tests and builds under Node 22 (PowerShell: `$env:PATH = "C:\Users\maido\AppData\Local\Programs\nodejs;" + $env:PATH`). Failures in SQLite history tests under Node 26 are version problems, not code bugs.

```bash
npm run typecheck          # server tsconfig + frontend/tsconfig.json
npm run build              # tsc -> dist/, esbuild pane bundle -> frontend/dist/taskpane.js, gateway bundle
npm test                   # vitest, mock providers only
npx vitest run tests/peer.test.ts          # one file
npx vitest run tests/peer.test.ts -t "name" # one test
npm run dev                # pane (esbuild watch) + bridge (tsx watch) from the working copy; 8765 if the installed bridge is stopped, else 8777
npm run pack:release       # build the release .tgz into release/
npm run build:msi          # thin per-user MSI (needs the WiX 5.0.2 CLI); payload only: node scripts/build-msi-payload.mjs
```

Validation before reporting done: `npm run typecheck`, `npm run build`, `npm test` (CI runs these on windows-latest with Node 22.12/22/24, then `pack:release`).

- Live provider tests (`*.live.test.ts`) are excluded unless `LIVE_PROVIDERS` is set; `npm run test:live` / `npm run smoke` hit real CLIs.
- `tests/support/hermetic.ts` gives every test run its own `OFFICE_PARTNER_DATA_DIR`; tests must never read the real `%LOCALAPPDATA%\Excel Partner`.
- `npm run dev` uses the real data folder, certificate and sign-ins. Office loads the pane from 8765 only, so `office-partner stop` first to test inside Office; a browser tab at `https://localhost:<port>/` works on either port.
- Do not run `codex login --device-auth` (or other login commands) to probe: it signs the owner out.

## Architecture

### Layers (enforced by `tests/architecture.test.ts`)

```
protocol/    zod schemas, tool input schemas, messages, events, system prompts. Imports nothing.
providers/   ALL provider logic. Imports protocol only.
shared/      what bridge and gateway both use: ProviderSessionManager, tool definitions/catalog, MCP server management, result conversion.
bridge/      adapter for the task pane: HTTP routes, WebSocket, turns, SQLite history, skills, files, peer routes.
gateway/     adapter for the Anthropic Messages API.
excel_mcp/   stdio MCP entry point the provider CLI spawns; imports protocol + shared.
frontend/    task pane (bundled by esbuild); imports protocol only.
bin/         install/start/update/doctor CLI (.mjs/.cjs); imports only other bin/ files.
```

The architecture test also fails if:
- any folder outside `providers/` compares, switches on or looks up a provider name (`=== "codex"`, `case "claude":`, `["copilot"]`, …). Behaviour differences belong in the adapter, exposed via `capabilities` or adapter methods.
- `shared/`, `bridge/`, `gateway/`, `excel_mcp/` contain a terminal CLI's display name ("Claude Code", "Codex CLI", …); those come from the provider's `static terminal`.
- `providers/` reaches `bin/` anywhere other than `providers/src/profiled.ts`.
- `bin/runtime-manifest.json` doesn't list every `PROVIDER_IDS` entry with `minVersion <= version`.

### Providers

`providers/src/types.ts` (`ProviderAdapter`) is the single contract: detect, list models/commands, start/send/cancel/close session, `capabilities`, `startLogin`/`loginAccounts`, `logout`, `verifyReady`, `install`/`uninstall`. `BaseCliAdapter` (`base.ts`) is the shared implementation; each provider lives in `providers/src/<id>/` (`adapter.ts` plus its own helpers). `registry.ts` creates adapters; `profiled.ts` wraps an adapter with the managed runtime profile from `bin/`. `mock.ts` is the adapter used by tests.

Managed CLIs: `bin/runtime-manifest.json` is the one data file of install facts: per-provider CLI versions, plus the pinned Node.js and SQLite binary (with SHA-256) the MSI uses (`installer/msi/`, `scripts/build-msi*.mjs`, `scripts/msi/`); `bin/runtime-store.cjs` / `runtime-manager.mjs` install them into runtime folders under the data folder (id depends only on provider/package/version/platform); that install copies and hashes nothing. One CLI per provider; `setup` (which `update` runs after npm replaces the package) brings the selected providers' CLIs to the versions the release names, and a failed CLI install never blocks a product update. `bin/*.mjs` are copied to a temp folder and run detached by the updater (`bridge/src/updater.ts`), so they must stay self-contained.

### Office tool call flow

1. Tool schemas: zod in `protocol/src/tools*.ts` (+ `registry.ts` for host availability), advertised JSON definitions in `shared/src/tool-definitions*.ts` / `catalog.ts`. Adding an Office tool touches both, plus the executor in `frontend/taskpane/executor/` (`index.ts` dispatches; `word/` and `powerpoint/` subfolders per host) and usually `frontend/taskpane/ui/tool-view.ts`.
2. The provider CLI spawns `excel_mcp` (stdio) with `EXCEL_BRIDGE_URL`/`EXCEL_SESSION_ID`/`EXCEL_SESSION_TOKEN`/`OFFICE_HOST`; `forwarder.ts` fetches `/internal/tools` and posts calls to the bridge.
3. `bridge/src/office-runtime.ts` validates, applies host/capability gates and Ask/Auto approval, then sends the call over the pane's WebSocket (`pane-connection.ts`) to the frontend executor.

External MCP servers (`shared/src/mcp-servers.ts`) are exposed as their native tools with namespaced names (`ext_<server>__<tool>`); never as generic list/call wrappers or via discovery prompts. Enabling/disabling refreshes the provider's tool registration while keeping the conversation.

### Peer / connected files

Panes for different open files can read each other (`read_connected_file`) or hand a turn to another file's chat (`send_to_connected_file`). Terminal agents reach the same through `excel_mcp/src/peer.ts` (`office-partner mcp-peer`: `list_office_files`, `ask_office_file`) calling the bridge's `/peer/*` routes (`bridge/src/routes/peer.ts`), authenticated with `peer-token.txt` in the data folder. `bridge/src/terminal-agents.ts` auto-registers that MCP server with installed CLIs (opt out: `OFFICE_PARTNER_TERMINAL_AGENTS=0`).

### Frontend

`frontend/taskpane/main.ts` registers `Office.onReady`; `chat/boot.ts` starts the pane. Shared chat state lives in `chat/state.ts`, turn handling in `chat/turn.ts`, and provider events in `chat/events.ts`. Every turn exit uses `setBusy(false)` to release request ownership and streaming/tool UI. `tests/chat-session-metadata.test.ts` imports the real event handler with mocked UI services. The pane reads provider differences from `capabilities` in detection results, never from provider names.

## Product rules (owner-confirmed)

- **Gateway is an API for Anthropic's official Claude Office add-ins**, not for this project's pane. The official client supplies tool schemas and executes Office tools; the gateway returns `tool_use` and resumes on `tool_result`, preserving client schemas, names, results and continuation for all three hosts. Never require Office Partner to be open, add a "Gateway binding" in its Settings, inject this project's Office tool catalog into official-client requests, execute official-client Office calls through this project's executors, or claim this project's backup/undo for them. External MCP tools hosted by the gateway may run server-side via the shared MCP runtime. Do not reintroduce the old pane-binding design. `gateway/src/engine.ts` and `bridge/src/turns.ts` stay separate.
- Neutral "Office Partner / Excel Partner" branding; never ship customer documents, company references, personal paths or account identifiers. The package bundles no skills or templates (user skills live in `%USERPROFILE%\.excel-partner\skills`).
- Provider credentials stay under each CLI's control; the add-in never stores an API key. Exception: where the CLI itself supports API-key sign-in (OpenCode), the pane may relay the key straight to that CLI's credential store, never to Office Partner's config, history, logs, usage or diagnostics. Every provider signs in from the pane.
- External folders are read only when the user grants access (`referenceDirs`).
- Don't commit unless asked.

## Environment notes

- Source files are CRLF (autocrlf). The Bash tool can mangle quoting inside heredocs; for multi-line helper scripts, write a file and run it with `node` (or `python`: Python 3.12 is installed). The `skills/odoo-test` skill is Python; test it from its own venv: `skills/odoo-test/.venv/Scripts/python.exe -m pytest` (run in `skills/odoo-test`).
- Installed product data: `%LOCALAPPDATA%\Excel Partner` (config, cert, logs, `runtime-state.json`). Stop only processes you started, by PID or port.
- `bin/doctor-checks.mjs` intentionally keeps its own per-provider auth probing (its tests inject a fake runner); known duplication, leave it.
