# Changelog

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
