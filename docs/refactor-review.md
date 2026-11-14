# Architecture refactor: independent review

Reviewed from the working tree on branch `main-office` by a session that did not write the refactor. Evidence is from running `npm run typecheck`, `npm run build`, `npx vitest run` (Node 22) and reading the code; nothing here relies on the handoff's own claims alone.

## Verdict

The five phases are done as the plan describes, with the exceptions in "Deliberate exceptions" below. Typecheck, build and the full suite pass. This review found gaps in how the rules were enforced and fixed them in the same pass (see "Fixed in the review").

## Phase checklist

| Phase | Result | Evidence |
|---|---|---|
| 0 Test baseline | Done | Hermetic data folder (`tests/support/hermetic.ts`), suite green under Node 22 |
| 1 Provider contract | Done | `ProviderAdapter` has `capabilities`, `startLogin`, `loginAccounts?`, `verifyReady`, `logout`, `install`, `uninstall`; Copilot declares `logout: false`; bridge login manager is a thin per-window map |
| 2 One CLI per provider | Done | `runtime-state.json` keeps `schemaVersion, selected, profiles, active, owners, disabled, failures`; no per-chat pinning, previous, pending or fallback profiles in `bin/` |
| 3 Dependency direction | Done, enforced | `tests/architecture.test.ts` (folder rules, `bin/` self-contained, provider-name rule, providers-to-bin rule, manifest rule) |
| 4 `frontend/` | Done | Renamed; no provider-name checks; `main.ts` is 4 lines and the chat logic lives in `frontend/taskpane/chat/` (largest file 258 lines) around one shared `app` state object |
| 5 Cleanup | Done | Unused schemas, error class and imports removed |

## Fixed in the review

1. The provider-name rule still scanned the old `addin` folder, so `frontend/` was not covered. It now scans `shared`, `bridge`, `gateway`, `frontend` and `protocol`, and also catches `"x" ?`, `includes("x")`, `map["x"]`, not only `===` and `case`.
2. That stricter rule found one more hard-coded default (`"copilot"` in six places). It is now `DEFAULT_PROVIDER` in `protocol/src/providers.ts`.
3. The provider id list was repeated in `protocol/src/events.ts` and `history.ts`; both now use `ProviderIdSchema`.
4. The minimum CLI version lived twice: `MIN_VERSION` in each adapter and the pin in `bin/runtime-manifest.json`; the doctor read the adapter source with a regex. The manifest is now the only source (`minVersion` per provider); adapters read it through `minimumVersion()`, the doctor reads the manifest, and a test checks every provider has one and none exceeds its pinned version.
5. `bridge`'s `SessionManager` (workbook sessions) shared a name with `shared`'s `ProviderSessionManager`. It is now `WorkbookSessions` in `bridge/src/workbook-sessions.ts`.
6. Stale `addin` names: `tsconfig.json` exclude and the `build:addin` script (now `build:frontend`).

## Deliberate exceptions (recorded, enforced by test where possible)

- `providers/src/profiled.ts` reaches `bin/` (`runtime-store.cjs`, `manage.mjs`, `runtime-manager.mjs`) by path to install and uninstall CLIs. `bin/` has to run alone when the updater copies it, so the install code stays there. The test allows this file only.
- `install()` / `uninstall()` are instance methods that call the manifest-driven tools, not static per-provider logic: nothing about installing differs between providers.
- `bridge` and `gateway` import `providers/` directly; `shared/` does not re-export it.
- `excel_mcp` may import `shared`.
- `bin/doctor-checks.mjs` keeps its own per-provider sign-in status commands (its tests inject a fake runner).

## Still open (needs the owner)

- `AGENTS.md` says the add-in never asks for an API key; OpenCode's account form does (the key goes to OpenCode, not stored).
- Nothing is committed.

## Closed after the owner decided

- `installer/install-office-partner.cmd` is kept (end-user installer).
- `main.ts` was split: `chat/state.ts` (the `app` object), `chat/services.ts` (the panels and widgets, built once), `chat/messages.ts`, `chat/composer.ts`, `chat/turn.ts`, `chat/events.ts`, `chat/chats.ts`, `chat/workbook.ts`, `chat/bridge-handlers.ts`, `chat/listeners.ts`, `chat/boot.ts`, plus `ui/elements.ts`. `tests/chat-session-metadata.test.ts` now runs the real `handleEvent`; `tests/pane-bundle.test.ts` runs the esbuild bundle so module initialisation order is covered.
- `logout()` was run for real against the installed OpenCode and Codex CLIs. Codex was already signed out (the command ran, nothing to lose). OpenCode: a made-up key for a provider the owner does not use was signed in and out again. That found a real bug: the sign-in service kept its own copy of the accounts, so the list still showed the account as connected after the credential was gone. `OpencodeAdapter.logout` now stops that service around the CLI call. Claude uses the system CLI and Copilot has no sign-out command, so neither was signed out.
