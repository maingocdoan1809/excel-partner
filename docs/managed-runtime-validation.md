# Managed runtime validation

Tested on Windows on 2026-11-10, on `feat/managed-provider-runtimes`, based on
`main-office` release 0.1.22. Source changes and test runs use a separate worktree.
The original checkout stays clean on `main-office`.

Final checks passed: `npm run typecheck`, `npm run build`, 1,793 tests across
141 files, and package dry-run inspection. The package includes the shrinkwrap,
management tools and validation documentation; local test data is excluded.

## Browser and real provider checks

Chrome was opened visibly at `https://localhost:8765/index.html`, served from the
feature worktree. The pre-existing installed bridge was idle with no connected
sessions; it was stopped for each check and restored afterward. Test settings,
provider distributions and provider configuration were isolated from user data.

The deterministic browser suite covers 16 flows: boot, chat, provider switching,
saved History, external MCP discovery/selection/approval/execution, navigation,
authenticated diagnostics and JSON download, temporary debug on/off, unavailable
rollback feedback, narrow recovery controls, both themes and browser errors.
Widths include 320, 380 and 1200 pixels. Screenshots and machine-readable results
are saved under the git-ignored `output/playwright/office-review` directory.

A separate live browser run uses the real private Copilot CLI 1.0.95 with SDK
1.0.18. The model receives the actual namespaced echo tool schema. Three tool
round trips are checked: the initial request, continuation in the same native
conversation, and reload followed by History resume with the recorded runtime
profile. Each request executes the harmless local MCP echo once after approval.
No fake provider is used in this run. Artifacts are under
`output/playwright/managed-browser-live`.

Real private CLI installation and native initialization checks cover:

| Provider | Exact CLI | Result |
| --- | --- | --- |
| Copilot | 1.0.95 | Version, integrity, sign-in, SDK and native registration passed; browser model/tool flow passed. |
| Claude | 2.1.294 | Installed and version checked; private sign-in required. Staged profile stayed pending. |
| Codex | 0.162.1 | Installed and version checked; private sign-in required. Staged profile stayed pending. |
| OpenCode | 1.19.5 | Version, integrity and native session/MCP registration passed. Paid-model access was not established. |

Source-keyed test caches avoid reusing an adapter snapshot after its manifest or
provider code changes. CLI sign-ins remain under provider control. No tokens are
copied, and no global provider package or PATH entry is changed.

## Fixes identified by testing

- A fresh Windows checkout converted the doctor's hashbang file to CRLF. Vite's
  transformed test module then contained a misplaced hashbang. `.gitattributes`
  preserves LF for JavaScript management scripts.
- The HTTPS browser fixture's in-process Node MCP client lacked the local CA.
  Its subprocess now receives that CA at startup, matching real MCP children.
- The old OpenCode 1.0.0 baseline does not support the adapter's question-server
  launch flags. The release pin is now the verified 1.19.5 distribution.
- OpenCode's older Windows batch launcher uses `start`, which is unsuitable for
  piped ACP. Managed installation can resolve its exact platform executable
  directly. Private OpenCode children also discard inherited `OPENCODE_BIN_PATH`.

## Coverage limits

Busy-service refusal, failed activation, interrupted update recovery, cached
offline rollback, checksum repair and owner-aware uninstall have automated
contract tests with fake installers and service lifecycle hooks. A production
Office installation was not upgraded, rolled back or uninstalled during these
checks. Browser rollback with no cached release was exercised directly.

Excel was not running and no connected Office document session was available.
These checks therefore do not prove Office.js execution or writes in a real
Excel workbook, Word document or PowerPoint deck. Claude/Codex model tool calls
still need sign-in to their private profiles before live validation. Remote
service behavior and organization policy remain outside local version isolation.

## Repeat the opt-in checks

After building the feature checkout:

```powershell
npm run typecheck
npm run build
npm test
npm run test:browser
$env:LIVE_MANAGED_RUNTIMES = 'copilot,claude,codex,opencode'
node scripts/test-managed-runtimes-live.mjs
```

The real provider/browser script requires an idle free port 8765 and an existing
localhost TLS certificate pair. Set `BROWSER_TEST_CERT`, `BROWSER_TEST_KEY` and
`LIVE_MANAGED_BROWSER=copilot`, then run
`node scripts/test-managed-browser-live.mjs`. It uses model tokens and opens a
visible browser. Restore any existing bridge after testing. A CLI needing sign-in
is reported as unverified; it is never counted as a successful model test.
