# Privacy and data flow

## What leaves the machine

- **To the AI provider the user selects:**
  - What it sends: prompts, plus the document content the model reads through tools (cells, ranges, slides, text) and attached files.
  - How: through that provider's own CLI (Claude Code, GitHub Copilot CLI, Codex, OpenCode), signed in with the user's own account.
  - Terms: the provider's terms, data handling and region apply, not Office Partner's.
- **Office Partner itself sends no telemetry or analytics.** It contains no analytics SDK.
- **Other connections:**
  - The npm registry (or the configured internal registry): update lookups, the package, provider CLIs.
  - nodejs.org: only the installer, when it adds a portable Node.js.
  - Playwright browser download: only when browser features are first used. Admins can turn these off with `allowBrowserAutomation: false`.
  - cdnjs.cloudflare.com and cdn.jsdelivr.net: scripts used inside artifact previews.
  - appsforoffice.microsoft.com: Office.js, loaded by Office.
- **Provider CLIs may send their own telemetry.** See each vendor's documentation. Office Partner turns off their self-update only.

## What stays on the machine

All local data is under `%LOCALAPPDATA%\Excel Partner`, readable only by the user (and administrators):

- **`history.db`:** chat history in SQLite, unencrypted, with full messages and tool results.
  - A daily backup is kept beside it.
  - Admins can set a retention period (`historyRetentionDays`).
- **Logs:** rotated. They record tool names and error codes, not tool argument values.
- **`usage.jsonl`:** token and request counts, size-capped.
- **Provider sign-ins:** kept by each CLI in its own state folder.
  - Office Partner never stores API keys.
  - When a provider signs in with an API key (OpenCode), the pane passes the key straight to that CLI's own store.

## User control

- **Ask mode is the default.** In Auto mode the assistant acts without asking, including running code and reading files outside the granted folders. Administrators can turn Auto off (`allowAutoApprove: false`).
- **External folders are read only after the user grants access.**
- **Administrators can turn off providers and risky features machine-wide.** See [admin-policy.md](admin-policy.md).
