# Agent Gateway

A standalone local Anthropic Messages-compatible API backed by the installed
Claude, GitHub Copilot, Codex and OpenCode CLIs. It reuses Excel Partner's
provider transports, authentication and cancellation code. It requires neither
the Excel Partner add-in nor its bridge. Provider CLI installation and sign-in
are still required.

## Install on Windows

The gateway ships in the same package as the Excel add-in (the add-in itself is
optional). Requirements: Node.js 20 or later, and at least one provider CLI
(Copilot, Claude Code, Codex or OpenCode) installed and signed in.

```powershell
npm install -g excel-partner
excel-partner setup gateway
```

Setup prints the URL and token and copies the token to the clipboard, so you can
paste it straight into Claude for Excel. Use `excel-partner setup all` to set up the
add-in and the gateway together. The `agent-gateway` command is the same CLI.

Setup creates a localhost certificate (or reuses the add-in's, so only one
certificate is trusted), trusts it for the current Windows user, generates a token,
registers a per-user Windows sign-in startup entry and starts an independent
background process on port 8766. It does not register an Office add-in or alter an
installed provider's credentials.

| Command | Description |
| --- | --- |
| `excel-partner setup gateway` | First-time setup: certificate, token, startup entry, start, warmup, then print URL and token. |
| `excel-partner info --copy` | Show the URL and token again and copy the token. `excel-partner token` prints only the token. |
| `excel-partner start gateway` / `stop gateway` / `restart gateway` | Control the background process. |
| `excel-partner status` | Show the add-in bridge and the gateway (add `gateway` for only the gateway). |
| `excel-partner doctor gateway` | Check Node, certificate, trust, startup entry, process, token and each provider, and say what to fix. |
| `excel-partner logs gateway` | Show recent log lines. |
| `excel-partner gateway models` | List the model IDs the gateway currently advertises. |
| `excel-partner gateway warmup` | Run one real minimal inference to fill the model verification cache. |
| `excel-partner gateway autostart on` / `off` | Add or remove the Windows sign-in startup entry. |
| `excel-partner update` | Install the latest release and restart the gateway (the whole package updates together). |
| `excel-partner uninstall gateway` | Stop the gateway and remove its startup entry. |

Settings commands (`config`, `alias`, `model-name`, `provider`) apply immediately:
the gateway restarts itself if it is running. Every command above also works as
`agent-gateway <command>`, for example `agent-gateway doctor`.

To remove the program, run `excel-partner uninstall all`, then
`npm uninstall -g excel-partner`. Configuration, certificate and token stay in
`%LOCALAPPDATA%\Agent Gateway`. Note that `excel-partner uninstall` (add-in) removes
trust for the shared certificate; run `excel-partner setup gateway` again if the
gateway stops being trusted afterwards.

Use `https://localhost:8766` and the token from `excel-partner info` in Claude for
Excel's Gateway sign-in. Select `copilot/default`, `claude/default`,
`codex/default`, `opencode/default`, or a discovered `provider/native-model-id`.
Only enabled providers appear. Native model discovery varies by CLI; a default
entry remains available when a CLI does not return a catalog. Do not invent a
native model ID: configure a tested alias or use `provider/default`.

Claude for Excel's current model picker filters IDs to names containing
`claude`, `sonnet`, `opus`, `haiku`, or `anthropic`. The gateway also publishes
`copilot/default [Claude API]`, `codex/default [Claude API]`, and corresponding
native model aliases for IDs that would otherwise be hidden. The suffix
describes API compatibility; the leading provider/native model is the actual
destination. `codex/gpt-6-sol [Claude API]` runs Codex with `gpt-6-sol`. Earlier
`anthropic/provider/model` aliases remain accepted but are no longer advertised.
Bare `claude-*` names are rejected with HTTP 400 unless a configured alias matches
them, so a request never silently runs on a different provider. The one exception
is Office's exact one-token `x` login probe, which is routed to the default
provider. Set `advertiseCompatibilityModels` to true to publish and accept
`claude-opus-4-7`, `claude-sonnet-4-6` and `claude-haiku-4-5-20251001` (routed to
the default provider or their alias); this is useful when a client only lets you
type a Claude model name.

To use friendly names rather than technical IDs in the Office picker:

```powershell
agent-gateway model-name set codex/gpt-6.1-sol "Codex — GPT-6.1-Sol"
agent-gateway model-name set copilot/default "Copilot — Default"
```

These names are stored in `config.json` under `modelNames`. The Office client
ignores `display_name`, so named aliases are used as public model IDs. A short
`(Claude API)` suffix is added when needed to pass its filter. Technical IDs and
previous protocol aliases continue to route as before but are not duplicated in
discovery. Names must be unique. Reconnect the Office client to reload its menu.

The current Office client also gives model verification requests only five
seconds. The gateway preloads model discovery and runs one real, minimal default
provider inference when you run `agent-gateway setup` or `agent-gateway warmup`,
then caches that exact one-token `x` probe for ten minutes. Only this precise probe
shape is cached; user conversation responses are not.
`x-agent-gateway-cache: verification-probe` identifies hits, which report zero new
inference usage. This is not a fabricated readiness response. Warmup consumes a
little provider quota, so it does not run at every start by default; set
`warmupOnStart` to true to run it at each start. Run `agent-gateway warmup` before
a later login if the verification cache expired and the provider starts too slowly. The server and API remain
available if warmup fails, but model verification can still fail legitimately.

Data lives in `%LOCALAPPDATA%\Agent Gateway`, separate from Excel Partner:
`config.json`, `token.txt`, `localhost.crt`, `localhost.key`, `server.json`,
`gateway.log`, and `gateway.err.log`. `AGENT_GATEWAY_HOME` overrides this folder.
The server binds only to IPv4/IPv6 loopback. Logs contain routing and HTTP status
metadata, not prompts, request bodies, tool results or authentication headers.

## Routing and configuration

```powershell
agent-gateway provider enable claude
agent-gateway provider enable codex
agent-gateway provider enable opencode
agent-gateway config set defaultRoute.provider copilot
agent-gateway config set defaultRoute.effort high
agent-gateway alias set quick copilot default
agent-gateway alias set claude-haiku-4-5-20251001 copilot default
agent-gateway config set requestTimeoutMs 180000
agent-gateway config set continuationTimeoutMs 600000
agent-gateway config set idleTimeoutMs 1800000
agent-gateway config set maxSessions 8
```

Explicit IDs (`copilot/native-id`) route to that provider and model. Configured
aliases take precedence. Unmapped `claude-*` IDs route through `defaultRoute`:
Claude for Excel issues helper requests using standard Claude IDs independently
of the selected chat model. `/v1/models` labels these mappings, and response
headers `x-agent-gateway-provider` / `x-agent-gateway-model` identify the actual
destination. A failed provider is never silently replaced with another provider.

`agent-gateway config show` displays all settings. Additional configurable values:
`port`, `maxBodyBytes`, `allowedOrigins`, `defaultRoute.model`, and
`enabledProviders`, `warmupOnStart`, `probeCacheTtlMs`, and
`advertiseCompatibilityModels`. Changes apply immediately. Remove an alias by
editing `config.json`; there are no credentials in this file.

## API and tool ownership

- `GET /v1/models`: discover routes; authenticated.
- `POST /v1/messages`: text and client tools, JSON or incremental SSE.
- `GET /health`: loopback health and active session count.
- `GET /gateway/status`: authenticated configuration/provider diagnostics.
- `POST /gateway/warmup`: refresh the real default-provider verification result.
- `POST /gateway/shutdown`: authenticated graceful stop, used by the CLI.

Authenticate with `x-api-key: <token>` or `Authorization: Bearer <token>`.
The default CORS origin is `https://pivot.claude.ai`. OPTIONS and error responses
carry CORS headers. The relay has a separate per-process internal token; the
public API token cannot call `/internal/*`.

Tools come from the API request, not from Excel Partner's workbook tool catalog.
The CLI sees them through a dedicated MCP relay. When a tool is selected, the
gateway returns `tool_use` and suspends the CLI handler. The client executes it
and sends `tool_result`; the handler is resolved and the same CLI turn continues.
Parallel tool calls are supported; `disable_parallel_tool_use` exposes one at a
time. Tool names/arguments/results are never executed locally by this relay.

Continuation IDs, original conversation history, system/tool configuration,
route and authentication identity must match. Tool results cannot be consumed
twice. Expired, mismatched or lost continuations return an error without replay.
Pending continuations remain in memory until completion, timeout or shutdown;
they do not survive a gateway restart.

The add-in and API use the same `providers/src/session-manager.ts` lifecycle.
A completed generation keeps its worker alive. Claude Office's
`metadata.user_id` contains a JSON string with its stable `session_id`
(`historyChatId`). The gateway hashes authentication identity, provider and that
session ID into the primary conversation key. Different IDs never share a worker,
even with identical transcripts. Device IDs and auxiliary `aux:...` inference IDs
are not chat identities. Clients without a recognized ID use full-history matching.

History checkpoints validate continuity and identify newly appended messages;
they no longer select Office chats. Cache hints and string/text-block
representation are normalized. Office's exact historical-image redaction marker
may replace an older image without changing session identity; actual text, tool
IDs and unredacted image edits still invalidate a match. Only newly appended messages go to the live
worker. Configuration changes reopen it through the shared manager and resume
the confirmed native conversation; a dead runtime follows the same recovery path.
Without a confirmed native ID, a replacement runtime receives the supplied full
history. A failed native resume is reported without silently retrying the prompt.

For an explicit ID, edited/truncated history, repeated completed requests and
simultaneous turns return 409; the gateway does not guess a new branch or replay
the conversation. An owned pending `context_snip` call may rebase earlier history
when the explicit ID and exact last assistant/tool-call response still match.
Other compaction that rewrites a saved prefix requires a new chat; arbitrary
rebasing is not supported. Without an ID, independent chats,
edited/truncated history and ambiguous matches open separate sessions. No request
is grouped by token or model alone. Idle workers expire after `idleTimeoutMs` (30 minutes by
default); `maxSessions` bounds retained workers and evicts the oldest idle worker
before rejecting a new request. Active/tool-waiting workers are never evicted for
capacity. Expiry, disconnect and shutdown close runtimes through the shared manager.
After clean idle expiry or capacity eviction, a bounded in-memory cache retains
confirmed native identities and history digests for up to 24 hours (at most
`4 * maxSessions` records). A matching follow-up resumes through the shared
session manager and sends only the appended messages. Shutdown clears this cache;
pending/failed turns are not archived or automatically replayed.

The standalone server also persists explicit-ID checkpoints to
`conversations.json` in the gateway data directory. It stores hashed keys, native
IDs, history digests and idle/inflight/blocked status, never API tokens, raw client
IDs, prompts, workbook values or image bytes. Completed chats can resume the same
native conversation after gateway restart. Storage uses atomic file replacement,
keeps at most 512 records and 32 MiB, and evicts older idle entries first. Corrupt
checkpoints stop startup rather than silently discard identity. Interrupted or
failed turns are blocked after restart: client tool actions may already have
executed. Pending tool continuations themselves are not restored. Do not run
multiple gateway processes sharing this data directory.

When cold reconstruction is required, historical images become explicit
unavailable-image markers with digests; base64 is never inserted into model text.
This does not restore visual access to those images. `maxPromptChars` (default
200000) bounds system plus transcript characters before inference; oversized text
is rejected rather than silently truncated. This is a character guard, not an
exact model token count.

## Compatibility boundaries

This is an agent-CLI adapter, not a raw inference proxy. Native agent instructions
and provider behavior can differ from the Anthropic API. Existing CLI tool and
sandbox restrictions apply; the gateway does not enable provider-native shell,
browser or filesystem tools. Caller tools are executed by the API client.

- Text, system text, conversation history, inline base64 JPEG/PNG/GIF/WebP image
  input, client tool schemas, `tool_use`, `tool_result` including errors, JSON
  responses and SSE are supported. Images are forwarded using each CLI's native
  image input format and are excluded from prompt text and gateway logs. Image
  URLs and document/file input are rejected; historical images cannot be restored
  after a native CLI session expires and is reconstructed from text history.
- `tool_choice: {type: "tool"}` exposes only the named client tool and asks the
  CLI to call it; `any` asks for at least one. CLI transports cannot enforce
  Anthropic's API-level forced-tool guarantee, so these choices remain best effort.
- `output_config.effort` maps to CLI effort; `max` maps to `high`. Thinking options
  are accepted as reasoning hints; native private reasoning/signatures are not
  exposed or fabricated. They do not guarantee identical Anthropic behavior.
- `max_tokens` is a best-effort response-length instruction. CLI transports do not
  expose an exact generation token ceiling or tokenizer. `/v1/messages/count_tokens`
  explicitly returns 501 rather than an invented count.
- Usage is reported when the underlying CLI turn completes. Intermediate tool
  responses contain zero counters because native whole-turn usage is not yet
  available; it is not charged again. If the CLI reports no usage, counters remain
  zero/unavailable. `x-agent-gateway-usage` describes this accounting behavior.
- Sampling controls, stop sequences and structured response formats are rejected.
  Uploaded-file APIs, Anthropic-hosted web/code-execution tools and prompt-cache
  accounting parity are not implemented.
- Tool configuration and routing must stay fixed through a pending tool round
  trip. Changes require a new generation/conversation.

## Claude for Excel customization researched

The official add-in supports more than a Gateway URL/token. Its deployment
configuration can supply custom inference headers, bootstrap sign-in/refresh,
MCP connections, feature restrictions, access policies and telemetry configuration.
These are settings of the official add-in, not provider CLI permissions.

Useful controls for this gateway include disabling Anthropic-hosted `web_search`
and `file.upload` when their unsupported server APIs would otherwise be used;
client-provided MCP tools can still be relayed. Permission restrictions such as
`permission_mode.accept_all_edits` can keep workbook edits in Ask mode.

Secrets in manifest URL configuration belong after `#`, not in query strings.
Prefer bootstrap configuration when distributing gateway credentials to users.
The current package implements gateway routing/effort/limits and CORS, not an
Entra/OAuth bootstrap service or changes to your official add-in manifest.

References:

- [Anthropic Office gateway requirements and configuration](https://claude.com/docs/office-agents/third-party-platforms)
- [Anthropic Office setup tooling](https://github.com/anthropics/financial-services/tree/main/claude-for-msft-365-install)
- [GitHub Copilot SDK](https://github.com/github/copilot-sdk)

## Troubleshooting

Run `excel-partner doctor gateway` first. It checks Node, the certificate and its
trust, the startup entry, the process, the token and every provider, and says what
to fix. `excel-partner logs gateway` shows recent log lines.

- **Office says it cannot connect, or shows a certificate error.** The gateway
  certificate is not trusted. Run `excel-partner setup gateway` again (it is safe to
  repeat), then restart Excel. Confirm the URL is `https://localhost:8766` and
  the token matches `excel-partner info`.
- **"Enter model name" / the model list cannot be retrieved.** Office waited
  five seconds and gave up. Run `excel-partner gateway models` to confirm discovery
  works, and `excel-partner gateway warmup` to fill the verification cache, then
  reconnect. Otherwise type `provider/default`, for example `copilot/default`.
- **"Recorded gateway is still alive but unhealthy".** `server.json` may point to
  a process id that Windows has since reused. Check that the id in
  `%LOCALAPPDATA%\Agent Gateway\server.json` is not a gateway process, delete that
  file, then run `excel-partner start gateway`.
- **"Port is occupied by another service".** Another program uses port 8766.
  Stop it, or run `excel-partner gateway config set port 8767` (the gateway
  restarts itself) and use the new port in Office.
- **A provider does not appear or fails with a login error.** Install and sign in
  to its CLI, then run `excel-partner gateway provider enable <provider>`.
- **409 "interrupted or failed turn" for a chat.** A turn that called tools and was
  then cancelled or failed may have changed the workbook, so the gateway refuses to
  replay it. Inspect the workbook, then start a new chat. A cancelled turn that
  never called a tool can simply be sent again.
- **PowerShell blocks `excel-partner`.** Run `excel-partner.cmd`, or allow scripts
  for your user with `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.
- **The gateway does not start at sign-in.** Run `excel-partner gateway autostart on`.

## Development and validation

Run the gateway from source with automatic restart. This uses a separate data
folder (`run/gateway-dev`) and port `8767`, and prints the URL and token:

```powershell
npm run dev:gateway
```

```powershell
npm run build:server
npx vitest run tests/agent-gateway.test.ts
npx tsx scripts/smoke-gateway.ts copilot
npx tsx scripts/smoke-gateway.ts claude
npx tsx scripts/smoke-gateway.ts codex
npx tsx scripts/smoke-gateway.ts opencode
node scripts/build-gateway.mjs
```

The smoke script uses real provider inference and can consume account quota. It
tests a synthetic read/write value in memory; it never modifies a real workbook.
All four installed providers passed this two-tool API round trip during initial
development. Set `CLAUDE_BIN` or the other provider binary overrides when needed.

## Latency diagnostics

`gateway.log` records `provider_session` actions (`runtime_opening`,
`runtime_opened`, `runtime_reused`) and `gateway_turn` phases. These include
session/turn IDs, history matching, worker setup time, first text, tool/client
waits and completion time. `afterTextMs` measures the tail after the last text
notification, including provider accounting/cleanup. No prompts, system text,
tool arguments/results or API tokens are included in these timing events.

Equivalent tool list ordering, default tool-choice fields, and string/text-block
system representations do not restart a worker. History matching also recognizes
adjacent text-block splits and string/text-block tool results. Actual content
changes and ambiguous histories still preserve conversation isolation.

## Client tool contract and provider checks

The gateway supplies the same API tool contract to the model session and MCP
initialization, so behavior does not depend on a CLI surfacing MCP instructions.
The relay advertises shared client-tool instructions from
`providers/src/client-tools.ts`. A client tool named `execute` or
`execute_office_js` means execution by the API client; it must not be substituted
with the provider's local interpreter. OpenCode V2 reaches MCP tools through
Code Mode (`tools.excel.<name>`); its adapter uses V2 permission and session-export
formats while keeping the V1 path for older installations. The standard ACP MCP
registration is retained rather than relying on unsupported ACP config fields.

Run `npx tsx scripts/smoke-gateway.ts <provider> --collision` for each enabled
provider. This uses real CLIs against synthetic client tools, including a remote
`execute` tool, nested success/result envelopes and an Office mock that executes
the generated code. Sheet3 already exists to test create-if-missing behavior.
The test asserts actual tool round trips, the written/read value, worker
reuse and memory on a second completed turn. It does not modify a real workbook.

Gateway tool-result conversion preserves the original text and error marker,
maps Anthropic inline base64 images to MCP image blocks, validates the result
against the MCP schema, and adds MCP `structuredContent` for a single JSON object.
Malformed or unsupported tool content becomes a valid MCP error; it never
automatically retries a workbook write. Logs distinguish client errors from
conversion errors without recording image bytes. This support concerns images
returned by client tools; it does not enable top-level image/document input.
Provider code tools
can inspect `success` and `result` without guessing whether the output is text.
Office.js usage guidance is attached only when an API client registers
`execute_office_js`; the Excel Partner MCP server/tool catalog is unchanged.
OpenCode V2 permission/export compatibility in this change is gated on explicit
API-supplied `mcp.toolNames`, leaving workbook add-in sessions on their existing path.

OpenCode retry diagnostics appear as `gateway_turn` / `provider_status` events
with `statusPhase`, attempt, HTTP status, error type and next retry time. They
exclude upstream headers, response bodies and free-form error messages. Older
OpenCode versions may not emit this metadata. Gateway timeouts remain bounded;
a timeout does not automatically replay the conversation or workbook writes.

The Office fixture accepts the async function body exactly like Claude for
Excel; it does not auto-invoke a declared `main`. Use `--office` to test the real
`execute_office_js` name/profile, or `--collision` to test a remote `execute` name.
Add `--line` with `--office` to check line creation and geometry as well.
Add `--redact-images` to `smoke-gateway.ts` to deliver a real image tool result,
replace it with Office's historical redaction marker, then assert reuse of the
same live worker and memory without replay. This checks session identity rather
than image interpretation.
Add `--expire-worker` to exercise native resume after idle expiry instead of
live worker reuse. Copilot API Auto omits reasoning effort because the CLI
rejects that parameter on Auto resume; explicit model and add-in settings keep
their existing behavior.
Add `--conversation-id` to test the Office metadata key, or `--restart-gateway`
to close the gateway and resume from its persisted native ID with only appended
input. These can be combined with `--redact-images`.

Run `npx tsx scripts/smoke-gateway-image.ts <provider> [native-model]` for a
real CLI vision check using two synthetic images and the same worker. This
consumes inference quota but never opens a workbook. Local checks passed for
Claude and Codex. Copilot returned incorrect colors despite successful MCP
delivery; an explicit `gpt-6-luna` passed one run but failed a later run, so
Copilot image interpretation is not considered reliable. The gateway does not
silently change models or infer image contents from text metadata.
