# QA harness

Excel Partner can turn a test plan kept in Excel into saved, repeatable Playwright
tests, run them, and write the results back to the workbook. This document is the
design; user-facing steps live in the README.

## Goals

1. **The AI writes tests; it never holds a key.** Passwords, cookies and tokens
   never reach the model, the workbook or git.
2. **One safe path to run anything.** Every script is validated, bound to an
   approval, and runs with a scrubbed environment and a network allowlist.
3. **Easy for a QA engineer.** A handful of commands and slash skills; no
   Playwright configuration by hand. Odoo (form login) is the first-class target.

Non-goals for now: CI integration, non-Chromium browsers, remote runners.

## Parts

```
Excel (TestPlan.xlsx) <-> task pane <-> bridge  --qa_* tools-->  QA core (qa/src)
                                                                    |
                                    excel-partner qa <cmd>  --------+
                                                                    v
                                  <project>/  qa.config.json
                                              automation/  tests/ _harness/ qa.manifest.json
                                              auth/  (gitignored)   results/ (gitignored)
```

* **QA core** (`qa/src`) is plain Node with no Playwright dependency. It owns the
  config, manifest, validator, secret vault, scaffolding and the runner process.
* **Project** is a folder the user creates with `excel-partner qa init`. Playwright
  is installed *into the project*, not into this package, so the package stays small
  and each project pins its own version.
* **`_harness/`** is a small runtime library copied into the project (fixtures,
  Odoo helpers, reporter, global setup). It is versioned and hash-checked. Tests
  may import only `@playwright/test` and `_harness` / `pages`.
* **Bridge tools** (`qa_*`) let the model save, read and run scripts and sync results.
  They are server tools, like `run_python`. Approval is deliberately *not* a tool.

## Project layout

```
qa.config.json                 environments, profiles, allowed origins, sheet columns
automation/
  qa.manifest.json             tcId -> file, status, sha256, approvedBy/At
  _harness/                    shipped runtime (hash-checked)
  pages/                       page objects written by the AI or the team
  tests/TC_001_login_ok.spec.js
auth/<profile>.json            Playwright storageState (gitignored, mode 0600)
results/<run-id>/              results.json, screenshots, traces (gitignored, pruned)
```

## Lifecycle and approval

```
          qa_save_script            qa approve TC_001 (human CLI only)
 (none) ---------------> draft ------------------------------> approved
                          |  ^                                   |
                 edit     |  | qa_save_script (any change)       | file hash changes
                          v  |                                   v
                        draft <------------------------------- draft (approval void)
                                                broken / flaky are set by the runner
```

* The manifest records the SHA-256 of the file at approval time. The runner
  re-hashes before every run; a changed file silently falls back to `draft`.
* `approved` scripts run unattended (`qa run`, `qa_run`). `draft` scripts run only
  supervised: headed by default, and the task pane asks the user every time, even
  in Auto mode.
* Approval exists only as `excel-partner qa approve`. It is not an MCP tool, so a
  model cannot approve its own work.

## Secrets

* Passwords are never written to the workbook, config or chat.
* `excel-partner qa login <profile>` opens a **headed** browser; the person types
  the password (and any OTP). The harness saves the resulting `storageState`.
* `excel-partner qa secret set <profile>` stores a password for unattended
  re-login, encrypted with Windows DPAPI (current user) under
  `%LOCALAPPDATA%\Excel Partner\qa-secrets\`. Elsewhere, environment variables
  `QA_<PROFILE>_USER` / `QA_<PROFILE>_PASSWORD` are read instead.
* Before each run the runner checks the session (`/web/session/get_session_info`
  for Odoo, or `profile.checkUrl`). An expired session is refreshed headlessly from
  the vault, or the run stops with the exact `qa login` command to run.
* The model only ever sees profile *names* and whether a session is valid.
* The authenticated exploration browser (`qa_open_session`) starts Playwright MCP
  with `--storage-state` and `--allowed-origins`, so the model can look at logged-in
  pages without seeing the credential.

## Guardrails

Validator (AST, `acorn`; applied on every save and again before every run):

| Rule | Why |
| --- | --- |
| Imports limited to `@playwright/test`, `../_harness/*`, `../pages/*` | no `fs`, `child_process`, network libraries |
| No `eval`, `Function`, `require`, dynamic `import()`, `process`, `globalThis`, `fetch`, `WebSocket`, `XMLHttpRequest` | no escape hatches |
| No `.route(`, `.newContext(`, `.launch(`, `.connect*(`, `.cookies(`, `.storageState(`, `.addCookies(` | cannot undo the allowlist or read tokens |
| No string literal that looks like a secret (JWT, `Bearer`, 24+ char base64) or a literal passed to a password-like field | no hard-coded credentials |
| Header `// TC: <id> | Req: <id>`; file name starts with the id; every `test(` title starts with `[<id>]`; at least one `expect` | traceability, no vacuous tests |
| File size and node count limits | keep scripts reviewable |

The validator is a guardrail against mistakes and prompt-injected scripts, not a
sandbox. Containment comes from the runner:

* **Scrubbed environment.** Only `PATH`, `SystemRoot`, `TEMP`, `USERPROFILE`,
  `LOCALAPPDATA`, … and `QA_*` reach Playwright. Provider tokens and the user's
  other variables do not.
* **Network allowlist, three layers.** (1) Chromium is launched with a dead proxy (`--proxy-server=http://127.0.0.1:9`) and a bypass list holding only the allowed origins, so every other destination, including IP addresses and other loopback ports, fails; (2) `--host-resolver-rules="MAP * ~NOTFOUND , EXCLUDE <host>"` makes DNS answer "not found" for other names; (3) the harness context routes every request through an origin check. Measured on Chromium 156: resolver rules alone let `http://127.0.0.1:<other port>` through, the dead proxy does not. `qa.config.json` without any origin refuses to run. Scripts cannot undo it: `page.route`, `newContext`, `launch` and the `request` fixture are rejected by the validator.
* **Dedicated profile.** Tests use a fresh context built from the profile's
  `storageState`; the user's own browser profile is never touched.
* **Evidence hygiene.** Screenshots on failure only; traces off unless requested;
  the reporter redacts cookie/`Authorization`/token patterns from error text;
  `results/` is pruned after 14 days.
* **Harness integrity.** `_harness/` hashes are in the manifest; a modified
  harness blocks the run until `qa upgrade` restores it.

## Excel contract

`TestCases` sheet columns (names are configurable in `qa.config.json`):

`TC_ID, Req_ID, Module, Title, Preconditions, Steps, Expected, Type, Priority,
Role, Automation_Status, Script_Path, Last_Run, Last_Result, Duration_s,
Run_Evidence, Notes`

Result sync is deterministic: the model calls `qa_sync_results`; the bridge reads
`results.json` and sends it to the task pane, which locates columns by header and
rows by `TC_ID` and writes them. The model does not copy results by hand.
`Requirements`, `RTM` (formulas), `Dashboard` and `Bugs` sheets come from the
generated `TestPlan.xlsx` template.

## Tools exposed to the model

| Tool | Does | Approval |
| --- | --- | --- |
| `qa_status` | environments, profile names + session validity, counts by status | none |
| `qa_open_session` | authenticated, origin-restricted exploration browser | Ask/Auto |
| `qa_list_scripts` / `qa_read_script` | read manifest / source | none |
| `qa_save_script` | validate, write `draft`, update manifest | Ask/Auto |
| `qa_run` | run approved (or draft, supervised) scripts | Ask/Auto, drafts always ask |
| `qa_results` / `qa_failure` | compact results; failure text + screenshot | none |
| `qa_sync_results` | write results to the sheet by `TC_ID` | follows workbook permission |

## CLI

```
excel-partner qa init [dir] [--preset odoo|generic] [--url <baseUrl>]
excel-partner qa login <profile> | secret set <profile> | secret clear <profile>
excel-partner qa validate [TC_ID...]     excel-partner qa list
excel-partner qa approve <TC_ID...>      excel-partner qa revoke <TC_ID...>
excel-partner qa run [TC_ID...] [--draft] [--headed] [--trace] [--tag <t>]
excel-partner qa doctor | upgrade | clean [--days 14] | use <dir>
```

`qa use <dir>` records the active project in `config.json` (`qaProject`). The
bridge may read and write only inside that project.

## What is verified, and what is not

Covered by tests (`tests/qa-*.test.ts`, about 150 cases): validator rules, approval hash and tamper detection, harness integrity, secret redaction, DPAPI round trip, session probe, CLI, bridge tools and approvals, sheet write-back, and an end-to-end run with real Playwright and Chromium against a fake Odoo (sign-in, pass/fail/flaky, evidence, allowlist with an image tag and fetch to a second origin, expired-session refresh, signed-in MCP browser).

Not covered, by design or for now:

* The validator is static analysis. A determined author can still find a way around it; approval by a person and the runner limits are the real controls. Scripts run with the Windows user's permissions. Node's permission model and a restricted token are possible hardening steps.
* Approval controls what the QA tools and the runner will execute. It cannot stop an assistant that you have allowed to run arbitrary code as yourself (`run_python` or skill scripts in Auto mode, `execute_office_js`): such code can edit `automation/qa.manifest.json` or the harness like any other file you own. Keep those tools on Ask, or turn them off, in projects where this matters.
* A corporate proxy in front of the staging site is not supported yet (the dead proxy replaces it).
* Password storage uses Windows DPAPI only; other systems use `QA_<PROFILE>_USER` / `QA_<PROFILE>_PASSWORD`.
* Dashboard and RTM formulas in `TestPlan.xlsx` were generated with openpyxl and checked for column references, not recalculated in Excel.
* No CI integration, one browser engine (Chromium).

## Failures and repair

A failed step is only fixable if the assistant can see why it failed. Two paths feed it that information:

* **Assistant drives the browser** (`playwright_browser`): a failed action returns the error plus a picture and the page's dialogs, invalid fields and messages; the same action failing twice, or five failures in a row, is refused until the next request (`bridge/src/browser-guard.ts`, `browser-diagnostics.ts`).
* **Saved script fails** (`qa_run`): the harness fixture records the page state at the moment of failure (`_harness/page-state.js`), and the reporter keeps it with the failing line, step and code. `qa_failure` returns them with a first-guess hint.

`/qa-automate` and `/qa-report` repair a failing script in the same turn (at most 2 rounds per test). Repair changes only locators, waits, set-up data and extra steps; a difference in behaviour is reported, never patched, and test-case text in the workbook is only ever proposed in Notes. A repaired script is a draft again and needs a person's approval; `run.healWithoutPrompt` only removes the per-run question in Auto mode.

## Long runs

A provider's tool call can time out long before a suite finishes, so `qa_run` starts the run in the background and answers within about 45 seconds: either the finished results or `state: running` plus a run id. `qa_results` then waits up to 45 seconds per call. One run per project at a time; Stop cancels it. Results are stored in `results/<run-id>/report.json`, so they survive a restart and `qa_sync_results` can be called later.

## Phases (all delivered)

1. QA core, CLI and unit tests (config, manifest, validator, vault, runner, scaffold).
2. Bridge tools, approval prompts, authenticated browser.
3. Excel template, `qa_sync_results`, installable skills, README.
4. End-to-end test with real Playwright against a local fake Odoo login.
