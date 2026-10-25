# Provider sessions

The add-in bridge and standalone API gateway use the same lifecycle for Claude, Copilot, Codex and OpenCode:

```text
Bridge HTTP/WebSocket handlers
    → ProviderSessionManager
        → ProviderAdapter
            → provider runtime + Excel MCP
```

`WorkbookSession` owns the pane connection, pending tool calls and the current
turn's abort signal. `ProviderSessionManager` owns runtime reuse, configuration
changes, native conversation identity and adapter cleanup. `ProviderRegistry`
selects adapters and caches discovery. Adapters implement their CLI protocol;
they do not decide whether a caller is switching chats. The shared manager accepts
`ProviderSessionState`, which has no workbook or socket dependency. The bridge's
original manager module re-exports it for existing callers.

The gateway maps full completed API history prefixes to retained conversations,
then calls the same `ensure`, `sendTurn`, `cancel` and `close` methods. It sends only
appended messages to a reused worker; model/effort/system/tools changes and dead
runtimes go through the shared configuration/native-resume rules. Ambiguous,
edited or missing history creates an isolated session. Pending tool continuations
remain strict and are never replayed.

## Lifecycle

The first prompt calls `ensure()` to open a runtime. Subsequent prompts reuse
that runtime. `sendTurn()` forwards events and holds `turn_finished` until the
adapter's iterator has completed cleanup, so the pane can immediately send its
next prompt.

| Transition | Runtime | Native conversation |
| --- | --- | --- |
| Next prompt, same settings | Reused | Unchanged |
| Model, effort, tools or system prompt changes | Closed and reopened | Resumed |
| Open a saved chat | Closed and reopened as needed | Explicit saved ID |
| New chat | Closed and reopened | New ID |
| Stop | Interrupted | Retained when the runtime supports it |
| Disconnect or provider switch | Closed | Saved pane metadata can resume it later |
| Turn throws or returns an error | Closed | Confirmed ID retained for the next user turn |

An error never causes automatic replay of a prompt: a tool might already have
changed the workbook. A failed resume is reported, never replaced silently with
a fresh conversation. An opening runtime that finishes after cancellation or
disconnect is disposed rather than attached to stale workbook state.

There are three different IDs: the workbook connection/session ID, the internal
adapter runtime ID (`providerSessionId`), and the provider's native conversation
ID (`providerNativeSessionId`). Only native IDs are used for resume and usage
accounting; accounting keys are namespaced by provider. A locally proposed Claude
ID is not published until the CLI confirms it.

## Adapter contract

- `startSession(context)` opens one persistent runtime using the supplied model,
  effort, instructions, MCP endpoint and optional native resume ID. Release all
  partially created resources if initialization fails.
  Report runtime exits or forced shutdown through `context.onRuntimeClosed` so
  the manager can resume before the next prompt, including after cancellation.
- `sendTurn(input)` uses that runtime, emits events for that request only, and
  finishes after a native turn result. It must release per-turn listeners and
  state before its iterator completes. A completed turn must not close the
  runtime or its input stream.
- `cancelTurn(id)` interrupts the active turn. The bridge aborts the turn signal
  and has a watchdog to close an unresponsive runtime.
- `closeSession(id)` releases the runtime, pending questions and temporary
  configuration. It must be safe to call again.
- `dispose()` closes all remaining runtimes during bridge shutdown.

Configuration changes go through the manager, rather than ad hoc restart logic
in HTTP handlers. The context factory passed to `ensure()` builds MCP settings
after the new capabilities have been applied.

## Transports

| Adapter | Persistent transport | Turn boundary |
| --- | --- | --- |
| Claude | `claude -p --input-format stream-json`; `ClaudeProcess` owns stdin/stdout and control requests | `result` record |
| Copilot | Copilot SDK over the local CLI; ACP for callers without the pane question callback | SDK completion / ACP prompt response |
| Codex | `codex app-server`, JSON-RPC | Native turn completion |
| OpenCode | `opencode acp`, JSON-RPC | ACP prompt response |

Claude sends an `initialize` control request when opening and an `interrupt`
control request on Stop. `--resume` is only passed when opening a saved
conversation, not for every prompt. Tool permissions, native user questions and
per-turn usage translation stay in the adapter; workbook approval and execution
stay in the tool gateway and task pane.

## Maintaining and adding providers

Implement `ProviderAdapter` and register it in `createDefaultRegistry()`. Keep
provider-specific messages inside the adapter and emit normalized `ProviderEvent`
records. Extend the shared lifecycle tests with the new provider ID, then add
transport tests for native initialization, streaming, questions, cancellation,
crashes, usage and resume failures.

Relevant coverage:

- `tests/provider-session-manager.test.ts`: shared lifecycle and cleanup contract.
- `tests/native-session-accounting.test.ts`: reconnect, native IDs and usage rows.
- `tests/claude-session-confirmation.test.ts`: process reuse, native ID
  confirmation, initialization failure, interruption and request isolation.
- `tests/claude.live.test.ts`: real installed Claude CLI against a fake local
  Anthropic API, including Excel MCP and consecutive prompts.

Build the server before the real CLI tests so the MCP process uses current code.
Set `CLAUDE_BIN` to a supported executable if the test runner's PATH resolves a
different CLI installation from the terminal.

## Adding or updating a provider

Implement `ProviderAdapter` in `providers/src` and register it in
`createDefaultRegistry`. Wire the supplied MCP endpoint and preserve explicit
client tool names; provider-native interpreters are not substitutes for client
tools. Translate only that CLI's transport, permissions and usage format in the
adapter. Runtime lifetime and native resume remain in `ProviderSessionManager`;
API history identity and tool continuations remain in `GatewayEngine`.

Add the provider ID to the protocol/config/UI catalogs where applicable. Run the
shared lifecycle contract tests and the gateway smoke in both normal and
`--collision` modes. New tool calling conventions belong in the shared
`client-tools` contract or the adapter, not in a client-specific gateway branch.
Latency events distinguish runtime setup, model text, client tool waits and
accounting so provider changes can be checked without logging workbook contents.
