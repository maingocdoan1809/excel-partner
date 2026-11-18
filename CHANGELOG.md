# Changelog

## Unreleased

- **Test kit: repair in the same run.** `python -m qakit.tools.rerun <ID>` runs a case of a run again inside that run and its new result replaces the old one (the earlier pictures are deleted, no new run folder, no extra entry in the run list); `rerun --pending` carries on with the cases a pause left untouched; the dashboard shows the run as running again. The skill tells the assistant to look through the learned knowledge and the references before repairing, and to use `rerun`, not `retry`, for a run in progress. The dashboard keeps the scroll position and open sections when it updates, shows each step's own pictures, and opens a picture in a viewer inside the page that walks through the case's pictures.
- **Test kit: repair the case at once, case by case.** A case that ends needing repair (BLOCKED, NOT EXECUTED, or a FAIL not declared with `t.defect`) is run once more immediately from a clean page; if it still needs repair the run pauses right there (`run_paused`, shown as "paused to fix"), the cases not reached stay PLANNED, and only the rest is run after the repair, not the whole suite. `t.reach` adds test data round by round until a boundary condition holds (tags overflowing their area, a list longer than a page) instead of leaving the case PENDING. `python -m qakit.tools.watch` stands guard over a run and returns the moment something needs repair. The skill now requires writing and trialling one case at a time and finishing each before the next.
- **Test kit: self-repair instead of BLOCKED.** `t.submit` clicks a save control, photographs the moment, reads the validation messages and fills the empty required fields with valid test values before trying again; `t.ensure` creates the record a testcase needs when the screen has none, instead of marking the case BLOCKED; `t.defect` is the only way to record "the application is wrong" and needs the list of causes in the test that were ruled out. A FAIL is no longer something a step can claim without that. New `references/repair.md` gives the assistant the decision ladder (missing data, refused save, missing element, environment, real defect) and what to do about a bug inside the kit. Fixed: an event field named `file` crashed every case (`emit_to`); events with a reserved field name are no longer dropped.
- **Test kit: sign in without a password, fewer questions, read the screen.** The environment form can use a saved sign-in (Settings > Testing > Sign-ins): the kit starts already signed in, including sites whose login is a session cookie, and an expired sign-in is a "sign in again" message, never a switch to an attached browser. "Ask me before creating test records" is off by default. Sending a workbook now also builds the testcase catalog, and bilingual column headers such as `ケース番号/ID` are recognised. When a step fails the kit reads the screen (validation messages, invalid or required-empty fields, open dialogs), saves a picture and shows it on the dashboard; `retry.attempt` retries only after the cause was fixed and stops with the real reason when it was not. Browser waits default to 8 s, not 30. The plan (cases, steps, expected results) and the assistant's phases show on the dashboard before anything runs. A red box that cannot be drawn on an empty list now boxes its container instead of failing the step.
- **Test kit in the add-in (Settings > Testing).** One folder holds everything (qa.config.json, hooks, kit code, testcase copy, run results, what was learned); you pick it once. The Testing screen saves the environment and account (environment-variable names only, never a password), installs the kit and its skills (`system-test`, `autotest-init`) and the Python libraries and browser, and shows the run live. The testcase workbook you have open is sent to the kit when you write about testing (unsaved edits included); status, result and date are written into its own rows while the run goes. `https://localhost:<port>/tests` is a read-only dashboard with runs, steps, evidence pictures and the notes the assistant has saved.
- **Test kit hardening:** a missing data check or snapshot is UNVERIFIED (exit 3), not clean; a case cannot be PASS while an expected step was never recorded; a highlight that cannot be drawn stops the step; a failed cleanup exits non-zero; a run closes against its saved plan, so a planned case without a result is counted NOT EXECUTED. A run stops by itself when the same error repeats `run.max_repeat` times (3 by default); the cases not run are marked NOT EXECUTED with the reason. Every step is written to the run's `events.jsonl`.
- **Thin MSI installer (about 8 MB, per user, no administrator rights).** `npm run build:msi` builds `release/office-partner-<version>-<arch>.msi` (x64 and ARM64) from `.github/workflows/msi.yml`. It carries the add-in and its pinned dependencies, with the Windows SQLite and koffi binaries already inside, so a computer needs no download from GitHub and no compiler for chat history. At install it downloads only the Node.js and Copilot CLI versions pinned in `bin/runtime-manifest.json` (new `node` and `native` entries, with SHA-256 for the Node.js zip and the SQLite binary). It always uses that Node.js, never one already on the computer, so the SQLite binary always matches. It installs where `npm install -g` would, so `office-partner update` keeps working. Install with `msiexec /i ... /qn` (properties `PROVIDERS=copilot`, `NOSIDELOAD=1`). Python, Git Bash, Chromium and the other providers are not part of it.
- **`trace_formula` (Excel 365, ExcelApi 1.12):** a cell's precedents and/or dependents, across sheets, up to 3 steps and 300 cells and ranges, with each cell's formula and value. Another open workbook can be asked to run it too.
- **`batch` (Excel):** run up to 20 cell and structure tools in order in one call. Every operation is validated before the first runs; each one then goes through the same path as a single call (limits, the pane's approval and overwrite guards, the Changes list). It stops at the first failure (or continues with `stopOnError: false`) and reports how far it got. Nothing is rolled back.
- **`analyze_range` (Excel):** profile a table's columns, or filter, group and aggregate its rows (count, sum, avg, min, max, distinct) over up to 500,000 cells, and get back a few lines instead of the rows. Dates are filtered as YYYY-MM-DD. Another open workbook can be asked to run it too.
- **Lighter tool list (`toolSchemas: "lazy"`, on by default; switch in Settings > Compact tool list or set `toolSchemas: "full"`).** Every tool is offered by name and one line; `get_tool_definition` returns the full definitions (several in one call), matches a name that is slightly wrong and lists every tool when nothing matches. A tool whose definition was not read is not run and the answer carries the definition; an argument the tool does not have is refused instead of being dropped; tools that take no arguments keep their definition. 68-79% fewer definition characters.
- **The usage ledger records, per turn, the calls, failures and result size of each tool and the size of the tool definitions offered.** `node scripts/usage-tools.mjs` summarises them. Names and sizes only.
- **The tool catalog has a size budget per host** (`tests/catalog-budget.test.ts`).
- **Fix:** the lockfile pinned `@github/copilot-sdk` 1.0.18 and `koffi` 3.2.1 while `package.json` asks for 1.0.34 and 3.3.6, so `npm ci` failed and the published shrinkwrap carried the old SDK. One lockfile (`npm-shrinkwrap.json`) remains.

## 0.1.23

### Security
- **Only the signed-in Windows user's own processes can connect to the local bridge and gateway.**
  - 127.0.0.1 is shared by every user on a machine (Remote Desktop, fast user switching).
  - `/health` no longer lists sessions.
- **Admins can turn off Auto mode** (`allowAutoApprove: false` in the policy file).
- **The task pane has a Content Security Policy.** The `execute_office_js` approval states what it can reach.
- **Network paths (UNC, WebDAV, device paths) are refused before any file access,** so they cannot leak Windows credentials.
- **Workbook reading caps the expanded size of the file** (zip bombs).
- **Other hardening:**
  - Logs no longer contain tool argument values.
  - The management token is not inherited by provider CLIs.
  - The bridge refuses to serve plain HTTP without its certificate.
  - Doctor no longer loads modules by bare name.

### Reliability
- **Chat history:**
  - Switching releases no longer fails history requests with `HISTORY_STORAGE_FAILED`: the bridge stops taking requests before it closes the database.
  - A daily backup is kept.
  - Retention can be set by policy.
- **Crashes:** the bridge and gateway log a crash, stop provider processes and exit with code 70. Provider CLIs left behind by a crash are stopped at the next start.
- **Turns and sessions:**
  - A turn with no activity for 20 minutes is ended (`TURN_IDLE_TIMEOUT`). There is no limit on total length.
  - A busy chat is never closed to make room for another workbook.
- **`usage.jsonl` is size-capped.**

### Install and update
- **Silent installer.** `install-office-partner.cmd /quiet /version x.y.z /uninstall` with distinct exit codes, user-context check, pinned portable Node.js and English messages. See docs/enterprise-deployment.md.
- **The localhost certificate renews itself** 60 days before it expires. Uninstalling the add-in keeps the trust the gateway still uses.
- **Admin policy file** (`%ProgramData%\Office Partner\policy.json`): turn off providers, Auto mode, code execution, MCP, web search, browser automation, Office.js and updates; pin a version; set the registry and history retention. See docs/admin-policy.md.
- **Update rules:**
  - Updates look up releases with npm, so the proxy and registry configuration apply.
  - Older and deprecated versions are refused unless asked for.
- **Locks left by a crash or restart no longer block update, repair or setup.**
- **`setup --no-sideload`** for manifests deployed from the Microsoft 365 admin center.

### Package
- **Dependencies:**
  - `@github/copilot-sdk` 1.0.34 (security fix); `koffi` is now a direct dependency.
  - Non-breaking `npm audit` fixes.
- **The package now includes THIRD_PARTY_NOTICES.md and this changelog.**
- **CI on Windows (Node 22.12, 22 and 24) and tagged releases published with npm provenance.**

### Known issues
- **Advisory GHSA-mrhx-c7xc-m3w5 for zod 3 (memory use on oversized collections)** is fixed only in zod 4.
  - Inputs reach zod only from authenticated processes of the same user.
  - Migrating to zod 4 is planned.
- **Saved browser sign-ins are stored unencrypted** in the user's data folder.
- **The gateway prints its token at setup.**
