# Admin policy

Administrators can restrict Office Partner on a machine with one JSON file:

`%ProgramData%\Office Partner\policy.json`

Only administrators can write to `%ProgramData%` by default, so users cannot loosen the policy. `OFFICE_PARTNER_POLICY_FILE` points to another file (testing).

- **No file:** no restrictions.
- **A file that cannot be read or is invalid** (bad JSON, an unknown key, a wrong type): every risky feature is off and the product does not update itself until the file is fixed. The bridge logs the reason. The schema is `protocol/src/policy.ts`.
- The bridge reads the policy when it starts (restart it after a change). History retention and update settings are read again each time they are used.

| Key | Default | Effect |
|---|---|---|
| `disabledProviders` | `[]` | Providers (`claude`, `copilot`, `codex`, `opencode`) the pane cannot switch to or sign in to. |
| `allowAutoApprove` | `true` | `false` disables the pane's Auto mode. Every change is then asked for. |
| `allowCodeExecution` | `true` | `false` refuses `run_python` and skill scripts. |
| `allowExternalMcp` | `true` | `false` starts and exposes no user-configured MCP servers. |
| `allowWebSearch` | `true` | `false` turns off provider web search and web fetch. |
| `allowBrowserAutomation` | `true` | `false` refuses browser automation (Playwright) and saved browser sign-ins. |
| `allowOfficeJs` | `true` | `false` removes `execute_office_js` (arbitrary Office.js in the pane). |
| `autoUpdate` | `true` | `false` refuses `office-partner update` and the pane's Update button (`UPDATE_DISABLED_BY_POLICY`). |
| `pinVersion` | none | Exact version every update installs, e.g. `"0.1.23"`. It may also be older than the installed one. Any other target is refused (`UPDATE_PINNED`). |
| `registry` | npm's own configuration | npm registry for update lookups and installs, e.g. an internal mirror. |
| `historyRetentionDays` | none | Delete chats older than this many days. Applied at start and daily. |

Whatever the policy says, only processes of the signed-in Windows user can connect to the local bridge and gateway.

Example: fixed version from an internal registry, no code execution, no Auto mode.

```json
{
  "pinVersion": "0.1.23",
  "registry": "https://npm.example.com/",
  "allowCodeExecution": false,
  "allowAutoApprove": false,
  "historyRetentionDays": 90
}
```
