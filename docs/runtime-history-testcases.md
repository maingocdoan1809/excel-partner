# Runtime recovery and durable chat history — test cases

Scope: Windows desktop; branch `feat/managed-provider-runtimes`. Test cases are
written before implementation. Automated tests use local fixtures and mock
providers. Real provider prompts and real Office writes remain opt-in.

## Acceptance matrix

| ID | Trigger / action | Expected result | Coverage / result |
| --- | --- | --- | --- |
| R01 | Begin gateway request, enter maintenance, then finish request body | No provider turn starts after maintenance succeeds | PASS — agent-gateway.test.ts (partial body) |
| R02 | A request is awaiting external MCP discovery during drain | No late provider turn starts | PASS — agent-gateway.test.ts (delayed discovery) |
| R03 | Active gateway turn or pending official-client tool result | Maintenance/shutdown refuses; continuation is preserved | PASS — agent-gateway.test.ts, system-controls.test.ts; live update refused a busy gateway before safe retry |
| R04 | Claude runtime exits during initialization; retained runtime initializes | Fallback session stays open; no prompt replay | PASS — profiled-adapter.test.ts (failed callback isolation) |
| R05 | Auth/policy failure, resumed chat failure, or failure after prompt | No runtime switch or automatic replay | PASS — profiled-adapter.test.ts, managed-runtimes.test.ts (no replay/auth/resume) |
| R06 | Cached rollback app is intact but CLI dependency is corrupt/missing | Refuse rollback before draining/stopping services | PASS — managed-runtimes.test.ts (before drain) |
| R07 | Startup integrity fallback selects retained app with broken runtime | Do not silently select an unusable recovery target | PASS — managed-runtimes.test.ts (damaged retained runtime) |
| R08 | Disable provider, continue current and saved chat | Existing profile can continue; a new chat is refused | PASS — system-controls.test.ts (current, saved and new chats) |
| R09 | Uninstall add-in then gateway, and reverse order | Last owner removes owned releases/runtimes; user data stays | PASS — managed-runtimes.test.ts (both owner orders) |
| R10 | Windows supplies `Path`/`PathExt` in a copied environment | Copilot resolves the installed CLI with case-insensitive variable names | PASS — copilot-native.test.ts; actual Windows reproducer fails before/passes after fix |
| R11 | Official Office client sends a large Unicode system prompt to Claude on Windows | CLI starts with a short command; prompt preserved byte-for-byte; temporary file removed on close/failure | PASS — large Unicode prompt, bounded argv, exact content, fresh resume file and failure/close cleanup; Windows CLI 208k probe initialized |
| U01 | Settings at 280/320/380 px in both themes | Recovery/history actions form aligned groups; fields and labels stay inside cards; no overlapping or clipped buttons | PASS — 20-flow Chromium suite; grouped actions and field bounds at 280/320/380 px, light/dark screenshots reviewed |
| U02 | Open Settings and child panels at narrow/desktop widths | Header is at most 52 px tall; Back remains reachable; subtitles do not consume a second row | PASS — header reduced from 90 px to 48 px; narrow/desktop header bounds and child-panel navigation checked |
| D01 | Gateway answers HTTP 200 with ready=false during initialization | Report identifies gateway/port and startup state; failures highlighted in summary; readiness failure preserved | PASS — doctor/system-panel tests distinguish initializing gateway and name failing checks; ready=false remains a failure |
| D02 | Office throws for window.confirm; delete old chats / repair / rollback | Inline confirmation works without native dialogs; cancel/close/change of range or version performs no action | PASS — inline Cancel/confirm, close-panel cancellation and changed age/version guards; native confirm stub throws |
| H01 | New chat, messages, tool groups, artifacts and changes | Durable SQLite round trip preserves content and native identity | PASS — history-store.test.ts (messages/tools/artifacts/changes/native ID) |
| H02 | Restart bridge / reopen database | Saved chats remain available | PASS — history-store.test.ts; native Windows DB reopen |
| H03 | Long text, >400 messages or >40 chats | History is not silently truncated/deleted | PASS — history-store.test.ts (46 chats, 405 messages, >3 MB text) |
| H04 | Open History with large database | Paginated summaries; transcript loaded only on demand | PASS — history-store.test.ts, history-panel.test.ts (paged reads) |
| H05 | Edit existing chat from two clients with same revision | Stale write rejected; newer data is preserved | PASS — two SQLite connections + stale pane + pending-copy recovery |
| H06 | Retry the same successful save after losing its HTTP response | Idempotent acknowledgement; no duplicates | PASS — history-store.test.ts, durable-history.test.ts (lost acknowledgement) |
| H07 | Delete/clear then replay a delayed save | Deleted history is not resurrected | PASS — history-store.test.ts, history-api.test.ts (tombstone/generation) |
| H08 | Excel client requests Word/PowerPoint data or sends forged host | Bridge enforces host scope from authenticated session | PASS — history-api.test.ts, history-store.test.ts, durable-history.test.ts |
| H09 | No token / invalid token / disconnected session | History API rejects access | PASS — history-api.test.ts (missing, invalid and disconnected tokens) |
| H10 | Import legacy localStorage history; repeat import | Preserve all hosts and metadata; no duplicate/replaced chats | PASS — history-store.test.ts + Chromium migration/reimport flow |
| H11 | Migration/import fails halfway | Transaction rolls back; source data remains intact | PASS — invalid batch + SQLite trigger failure mid-import; source retained |
| H12 | Storage is full/blocked or legacy JSON is corrupt | Visible recoverable error; never replace history with empty/newest-only data | PASS — durable-history.test.ts, sessions.test.ts; blocked storage/corrupt JSON |
| H13 | Bridge offline, unsaved changes, reconnect | Queue survives pane reload when browser storage is available; retry after reconnect | PASS — durable-history.test.ts + Chromium close/reopen/outbox flow |
| H14 | Continuous streaming / switching chat / finishing turn | Bounded write interval; serialized saves; correct conversation ownership | PASS — history-scheduler.test.ts (continuous 500 ms checkpoints), serialized saves |
| H15 | Close/reload pane with pending save | Durable pending queue; best-effort keepalive flush, not pagehide-only persistence | PASS — Chromium close/reopen recovery + durable outbox; pagehide flush best effort |
| H16 | Open New chat after failed save | Previous unsaved chat is retained and clearly reported | PASS — durable-history.test.ts (failed chat retained while another chat saves) |
| H17 | Rename native conversation / bind provisional identity | Single history item per native conversation; old transcript remains correct | PASS — durable-history.test.ts (provisional → native), sessions.test.ts |
| H18 | Disk full, read-only path or SQLite write failure | Existing records retained; save fails explicitly; no false saved status | PASS — SQLite write failure injected with a trigger; no destructive real-disk fill |
| H19 | Corrupt DB / unsupported newer schema | Fail closed, preserve file, show diagnostic; never recreate over it | PASS — corrupt bytes/future version retained; history API 503, health 200 |
| H20 | History export | Full JSON export; credentials excluded | PASS — history-store.test.ts + Chromium export (the separate SQLite backup button was removed; export covers it) |
| H21 | User requests retention cleanup | Explicit scope/cutoff; never delete current/pending chats automatically | PASS — protected/unprotected cutoff cleanup + Chromium Settings flow |
| H22 | History removed / deleted in Settings | UI and count update only after server acknowledgement; failure visible | PASS — history-panel.test.ts (delayed/failed ACK), pane-boot.test.ts (delete/clear) |
| H23 | App update/rollback/full uninstall | Database remains outside immutable releases and is not removed; old localStorage retained for legacy rollback | PASS — managed-runtimes.test.ts retains history.db; browser source retained; see rollback limit below |
| H24 | Restore after managed pane reload with multiple open panes | Restore exact chat ID for that pane, not merely globally newest | PASS — pane-boot.test.ts (exact ID with newer chat in another pane) |
| H25 | Clean Windows install with supported Node versions | SQLite driver installs/loads without compiler/build tools | PASS on existing Windows x64 — Node 20.19/22.12/24.22 prebuilt load; pristine VM/ARM64 not run |
| H26 | History storage size/statistics and cleanup | Display database usage/count; paging, export and cleanup available | PASS — history-store.test.ts + Chromium Settings controls/statistics |
| H27 | Disconnect while loading History or opening a transcript | Visible retryable error; never misleading empty history | PASS — history-panel.test.ts (list/transcript failure and retry), durable-history.test.ts |
| H28 | Old localStorage quota test | Regression verifies old records survive failed writes; warning is observable | PASS — sessions.test.ts (quota throws; earlier browser bytes retained) |
| H29 | Continue a migrated legacy chat before its first revision is loaded | Safely rebase compatible transcript; no manual export or duplicate | PASS — durable-history + real SQLite integration |
| H30 | Two panes replay an older outbox after a newer commit | Discard already-contained snapshot; retain all newer messages | PASS — real SQLite + two-tab Chromium regression |
| H31 | Pending save fails while saved History is readable | Keep History available; pending failure never becomes global unavailable | PASS — connection, deleted-chat and offline cases |
| H32 | Two versions genuinely diverge | Save one stable separate copy; preserve original and explicit branch selection | PASS — live SQLite revision integration, continued edits and reload |
| H33 | Duplicated pane inherits sessionStorage writer ID | Each running pane keeps a separate outbox; both offline chats survive | PASS — duplicate-writer regression |
| H34 | Stream/tool completion arrives while follow-up is queued | Recognize monotonic updates; no unnecessary saved copies | PASS — real SQLite streaming integration |
| H35 | Temporary write failure with healthy bridge connection | Retry automatically with backoff; no Settings/export step required | PASS — timed offline/retry regression |
| O01 | Multiple panes discover OpenCode models concurrently or runtime changes mid-discovery | One in-flight discovery per provider/revision; stale results cannot replace current cache; failed discovery can retry | PASS — provider-model-catalog.test.ts (single flight, revision race, retry) |
| O02 | OpenCode retries after upstream overload | Show safe retry status for this session; old ACP uses authenticated status polling; other sessions and raw errors remain private | PASS — opencode-status/opencode tests; live Windows overload reproduced |
| O03 | Requested model selection fails | Return the selection error; never silently run another model | PASS — provider-native-resume.test.ts |
| O04 | ACP ends without assistant text or a tool call | Report empty_response; do not mark the turn completed | PASS — provider-native-resume.test.ts; browser progress/error flow |
| O05 | No initial model response for 120 seconds, or user presses Stop | Release composer and stop own CLI; streaming output/tools/approval can continue beyond initial deadline | PASS — fake-clock silent/streaming tests and cancellation regression |
| F01 | Repository validation | Typecheck + build + full default tests pass | PASS — typecheck, build, full default suite (see execution record) |
| F02 | Real Chromium UI regression suite | Chat, reload/resume, migration, History, storage actions, both themes/narrow pane | PASS — 20 real Chromium flows with mock providers |
| F03 | Package dry run | DB/WAL, fixtures, personal files excluded; SQLite dependency and management tools included | PASS — npm pack inspection; no user data or test fixtures included |
| F04 | Real Word/Excel/PowerPoint documents | Manual opt-in checks; report separately from mocks | NOT RUN — real Office document editing requires desktop Office validation |
| F05 | Real authenticated provider/model | Opt-in check only; report separately from mocks | PASS — real Copilot: text reply, same native ID after reload/resume, remembered nonce; zero tool calls |

## Persistence rules fixed before implementation

- The bridge owns `history.db` under application data, outside retained releases.
- Existing `localStorage` is migration input and rollback recovery, never the
  authoritative post-migration store. No destructive automatic quota fallback.
- Writes have a revision and idempotency key; deletes have tombstones/generations.
- Storage errors are surfaced. A rejected or unacknowledged write is not saved.
- UI keeps a bounded durable outbox; inability to queue is visible. No claim of
  guaranteed recovery after a crash if both bridge and browser storage fail.
- Retention is explicit. Database storage, rendered pages and model context have
  separate limits. Office undo stays memory-only; stored change records do not
  make prior document edits undoable.
- Old app releases may read their retained localStorage snapshot; schema upgrades
  must be additive/versioned and fail closed for unsupported future versions.

## Final execution record

Executed on 2026-11-11, on `feat/managed-provider-runtimes` (base `8a89c07`
plus the local fixes). The matrix was written before the implementation; R10 was
added when the installed Windows run exposed an additional case.

| Check | Result |
| --- | --- |
| `npm run typecheck` | PASS |
| `npm run build` | PASS — bridge, task pane and standalone gateway |
| `npm test` | PASS — 146 files; 1,839 passed, 2 skipped (1,841 total) |
| `node scripts/test-browser.mjs` with Chromium | PASS — 20 flows; zero uncaught page errors or failed requests |
| `npm pack --ignore-scripts --json` after build | PASS — 167 package files; SQLite code/driver dependency included; no DB/WAL/backups, fixtures, `.archify`, screenshots or personal presentation |
| Native Windows x64 SQLite | PASS — Node 20.19.0, 22.12.0 and 24.22.1: commit, reopen and revision conflict, using prebuilt driver 12.9.0 |
| Install into existing Windows installation | PASS — package installed; baseline retained; managed update committed with previous release available; bridge 8765 and gateway 8766 healthy |
| Installed Windows Chrome at `https://localhost:8765/index.html` | PASS — 6 checks: trusted HTTPS, authenticated history API/anonymous rejection, recovery controls, commit/conflict, reload/read, fixture-only cleanup |
| Installed Windows provider integration | PASS — real Copilot; 3 checks: response, same-ID reload/resume with remembered nonce, fixture-only cleanup; zero tool calls/page errors |

The two default-suite skips are Windows-only executable resolution tests in
`copilot-native.test.ts` and `managed-runtimes.test.ts`; they are not hidden live
provider checks. Default `*.live.test.ts` suites are excluded without
`LIVE_PROVIDERS`. SQLite tests on Windows run as actual Windows processes; the
full default suite and 19-flow Chromium suite run on Linux with mocks.

The installed update refused maintenance while the gateway was busy, retained
the running release, and committed after idle on a safe retry. No active Office
turn or tool continuation was forcibly interrupted.

Before-fix runs recorded failures for R01, R04, R06, R08 and R09. The new R10
regression also failed before its fix. On the installed Windows build, the
inherited environment had `Path`, and CLI resolution failed; explicitly
normalizing it to `PATH` made that same executable resolve. After the fix,
the unmodified Windows environment resolves the CLI directly. The final installed
Windows run passes both the 6-check storage/UI harness and the 3-check live
Copilot harness. The initial real-model run covered Copilot; the follow-up below adds Claude
Haiku. Codex and OpenCode remain covered by mocks in the default suite.

Limitations are explicit: no pristine Windows VM or ARM64 execution, no physical
disk exhaustion (write failure is injected transactionally), and no real
Excel/Word/PowerPoint document edits from the standalone browser page. Releases
predating SQLite see the retained browser migration snapshot, not newer
SQLite-only chats. This build supports Node 20.19+ in major 20, or 22.12+ below
major 26. No database service or SQLite installer is required.

Repeatable Windows storage validation is available as
`node scripts/test-history-windows.mjs` after building in a Windows checkout.
Temporary live-install browser harnesses/results are kept outside the package;
their synthetic chats are deleted explicitly, without clearing user history.

## Installed-Windows follow-up: Claude startup and narrow Settings

The official Claude for Office report at 2026-11-11T02:38:07Z exposed R11.
Claude instructions were passed as an inline `--system-prompt` argument. A real
Windows reproduction with 208,084 command characters fails with
`ENAMETOOLONG`; the file-backed form initializes successfully with a 208,000
character Unicode prompt and a 233-character command, without a model call.
The adapter now writes exact UTF-8 instructions into its private session
directory and passes `--system-prompt-file`. Existing session cleanup removes
the file on initialization failure and close. No truncation or prompt replay is
introduced. Mock regression also verifies fresh instructions on native resume.

The narrow Settings regression failed before styling changes. Provider cards,
diagnostics, restore controls and history recovery now have separate groups.
Buttons have consistent 38 px minimum height and aligned one-column layouts
at 280/320 px; 380 px uses a responsive grid. Both themes are captured and
reviewed, including the storage controls below the initial viewport.

Follow-up automated checks: typecheck/build pass, 146 test files and 1,837 tests
pass (2 Windows-only skips), and all 20 Chromium flows pass. Installed Windows validation passes: `claude/haiku` handles a 42,071-character
system prompt and a continuation, with zero tool calls. Native Chrome checks
pass at 280/320/380 px in both themes, with zero page errors.

The subsequent D01 report contained a real `SERVICE_NOT_READY` failure while
the gateway was initializing. Rechecking the installed service later returned
`ready: true`; both service checks passed. Overall status was `warn` only because
the private Codex runtime awaits sign-in. Diagnostics now name the service and
port, distinguish initializing from stopped/unreachable, and show failed check
IDs in the summary. The fail severity for an unready service is preserved.

U02 reduces the Settings header from 90 px to 48 px, uses a 16 px title and
28 px Back button, removes the subtitle row and applies compact spacing to
child panels. The final default suite passes 1,839 tests with two Windows-only
skips; browser validation includes headers, action grids and child navigation.

Final installed-Windows verification after U02/D01: all six width/theme
combinations measure a 48 px Settings header. The real UI Check system action
reports zero failures, identifies `gateway.health` and port 8766, and returns
`warn` only for the pending Codex sign-in. The final installed release also
passes the real Claude Haiku long-system-prompt request and continuation.

## History reliability and Office confirmation follow-up

Three pre-fix regressions reproduced the existing-user failures: continuing an
imported chat, replaying a superseded outbox on reopen, and treating a write
failure as a failed History connection. Writes still use SQLite CAS. The pane
now compares stable transcript content after a revision conflict: contained old
snapshots are acknowledged locally, monotonic extensions are retried against
the latest revision, and divergent content is saved as a stable separate copy.
Copies clear native resume identity and do not send or replay any provider turn.
Pane restoration follows its own copy; explicitly selecting the original row
selects the original branch. Deleting a copy removes its active alias.

Outboxes are unique per running instance even when sessionStorage is copied.
Temporary write failures retry with bounded backoff. A failed write does not
block reading saved History. Explicit deletion/generation conflicts are not
resurrected automatically; pending content remains available for recovery.
A consistent Windows SQLite backup was created before deploying these changes.

Both remaining native confirmation calls (old-chat cleanup and repair/rollback)
are replaced by inline controls. Cancel, Escape and closing Settings do not
perform the operation. Changing the age range or selected version requires a
new confirmation. Tests make window.confirm throw exactly as the Office host
does. Browser validation confirms cleanup without a native dialog.

## OpenCode response reliability follow-up

Real Windows OpenCode 1.19.3 accepted the existing ACP hostname/port arguments.
A minimal native request completed in 13.4 seconds; subsequent scratch/catalog/
custom-instruction controls completed in 23–33 seconds. The installed pane also
reproduced a silent wait beyond 75 seconds and a completed turn without text.
Native logs recorded `AI_APICallError: The inference service is overloaded`
for `opencode/big-pickle`. This evidence identifies upstream overload, without
claiming all delays come from that cause or changing the user's model/config.

O01–O05 were added before the fixes. Model discovery is shared per runtime
revision. Model-selection errors propagate. ACP retry metadata reaches the pane;
older ACP versions additionally use the same process's authenticated status API,
restricted to its session and sanitized retry fields. A silent initial response
is limited to 120 seconds; received text/thoughts, tool calls or permission
requests end that initial deadline. Stop releases a silent turn immediately.
An empty completed response is an explicit error, preventing automatic follow-up
submission after a failed response. No automatic model switch or prompt replay.

Validation after these changes: `npm run typecheck` and `npm run build` pass;
`npm test` passes 150 files / 1,861 tests, with two Windows-only skips. All 22
Chromium flows pass, including retry/error labels and composer recovery. Package
inspection excludes databases, user documents, test fixtures and architecture
artifacts; lockfile and shrinkwrap match.

The preceding history/confirmation release also passed four installed Windows
checks: stale shared outbox recovery across two panes, stable divergent copy with
original preserved, inline Cancel, and inline Confirm with native confirm forced
to throw. Cleanup confirmation used an intercepted response; only the test's own
history fixtures were deleted.

Installed Windows verification after the OpenCode follow-up:

- Active release `73465d75cc4d46d40caeb5f2c1c107c473bffa909ccc92ac09b3404ae0dc6f81`
  is committed. Global package matches the active release across eight checked
  provider, protocol, UI and diagnostic files.
- Both service health checks pass. Doctor has no failures; its remaining warning
  is the private Codex runtime awaiting authentication.
- All four installed history/confirmation checks pass again.
- Real OpenCode pane shows retry progress and releases the composer on the
  120-second initial-response timeout. A subsequent text-only test completed in
  115.3 seconds and resumed the same native conversation after reload in 24.2
  seconds, preserving its test nonce and executing no tools. Only test history
  was removed. These results do NOT certify stable latency.
- Controlled old/new adapter requests with the same CLI, scratch directory,
  tool schemas and instructions completed in 75.7 / 19.5 seconds respectively;
  one pair does not establish a performance fix. Native logs also recorded an
  overloaded inference service during the installed pane's timed-out attempt.
- Deployment encountered graceful service shutdown timeouts. Recovery required
  cleaning an idle recorded bridge process after verifying its executable and
  command identity, then activating the already verified prepared release.
  The underlying shutdown delay remains a follow-up; no active user turn was
  force-stopped for deployment.

User deferred further OpenCode investigation. They report previously selecting
an explicit model; the installed probes above used Default, so the exact former
model/configuration has not yet been compared. OpenCode remains an open issue.
