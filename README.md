# Excel Partner

Chat with Claude Code, GitHub Copilot CLI, or Codex CLI to read and edit the workbook open in Excel desktop on Windows.

Excel Partner installs a local bridge and an Excel task pane. Workbook changes run through Office.js in Excel. The bridge runs on your computer; each AI provider uses its own CLI and account.

## Requirements

- Windows with Excel desktop supporting ExcelApi 1.9 or later. Some features require newer Excel APIs.
- Node.js 20 or later and npm, installed on Windows.
- At least one supported provider CLI installed and signed in on Windows.
- Internet access and an account with access to the selected provider's service.

Run the commands below in Windows PowerShell. The npm CLI supports Windows desktop Excel. Provider CLIs installed only inside WSL are not available to this installation.

## Installation

### 1. Check Node.js and npm

```powershell
node --version
npm --version
```

If either command is unavailable, install Node.js with npm and reopen PowerShell.

### 2. Install a provider

Choose at least one provider:

| Provider | Install | Sign in |
| --- | --- | --- |
| GitHub Copilot | `npm install -g @github/copilot` | `copilot login` |
| Claude Code | `npm install -g @anthropic-ai/claude-code` | `claude` and follow the login prompts |
| Codex | `npm install -g @openai/codex` | `codex login` |
| OpenCode | `npm install -g opencode-ai` | `opencode auth login` |

Provider accounts, subscriptions, usage limits, and charges are managed by the provider. Provider CLIs are installed separately from Excel Partner.

### 3. Install and configure Excel Partner

Save your work and close Excel, then run:

```powershell
npm install -g excel-partner
excel-partner setup
excel-partner status
```

`setup` creates and trusts a localhost certificate for the current Windows user, registers the add-in, starts the bridge, and adds a Windows sign-in startup entry. Complete the certificate trust prompt if Windows displays one.

A healthy bridge reports `Bridge running at https://localhost:8765`.

To install a release tarball instead of downloading from npm:

```powershell
npm install -g C:\path\to\excel-partner-0.1.8.tgz
excel-partner setup
```

### 4. Open the add-in

1. Reopen Excel and open a workbook.
2. Select **Home > Excel Partner**.
3. Choose your provider in the task pane and complete sign-in if requested.
4. Select a cell range, enter a request, and send it.

For example: **Add a Total column in D. Calculate B times C for each row and bold the header.**

GitHub Copilot also supports device-code sign-in through **Settings > Sign in to GitHub Copilot**. Claude Code and Codex use their CLI login flows.

## Features

- Read selections, ranges, formulas, formatting, and workbook objects.
- Write formulas and values; format, sort, filter, and organize worksheets.
- Work with tables, charts, pivot tables, and shapes where supported by Excel.
- Inspect grouped flowcharts and connector attachments; use `validate_flowchart` to flag straight connector crossings, overlaps, and disconnected shapes. Routed connectors require visual review.
- Choose a provider, model, and reasoning effort from the task pane.
- Attach local files or browse workspace files with `@`.
- Add reusable instructions through **Settings > Manage skills**, then invoke them with `/skill-name`.
- Edit installed skill text files through **Manage skills > Edit**, including instructions, references and scripts. Binary templates and images are preserved. Saving rejects a stale revision instead of overwriting newer edits.
- Let the assistant update an installed skill with `read_skill` (`forEditing: true`) and `update_skill`. Ask/Auto applies to each save; updating a script does not run it.
- Send another message while the assistant works. It appears as **Pending** and reaches the assistant with the next tool result, or starts the next turn if no tool follows. Stop and disconnect preserve unsent messages for manual resend.
- Add optional shared instructions through **Settings > System prompt**. The shipped default is empty: no add-in system prompt is injected. Clear the editor and save, or choose **Restore default**, to remove custom instructions. Changes apply on the next request; start a new chat to avoid old instructions retained in conversation history. Provider-native instructions and tool permissions still apply.
- Review workbook changes and return to conversations through chat history.

Press **Enter** to send a message and **Ctrl+Enter** to insert a new line.

Skill ZIP uploads retain scripts, templates and other assets, except nested `.zip` files. The existing upload size and path limits still apply. Python (`.py`) and Bash (`.sh`) scripts can be inspected with `read_skill` and run through `run_skill_script`. **Ask** requests approval for each run; **Auto** approves execution automatically. Scripts run with your local user permissions on the bridge machine and can modify local files. The skill folder is the working directory; the runner passes arguments literally, supports text stdin, and returns stdout, stderr and the exit status. Runs default to 30 seconds (maximum 60 seconds) and stop when their combined output exceeds 512 KB, the turn is cancelled, or the pane disconnects. Optional `outputFiles` paths inside the skill folder return generated files as `attachment://` references, up to 30 MB combined. Use the Excel tools to modify the open workbook.

Install Python and any skill dependencies on the machine running the bridge. The runner uses `python` on Windows, `python3` on Linux/WSL, and `bash` for Shell scripts. On Windows it first checks common per-user and system Git Bash installation folders. Set `BRIDGE_PYTHON_BIN` or `BRIDGE_BASH_BIN` to an interpreter executable path when needed, then restart the bridge. Bash runs in the bridge's OS; a Windows bridge does not automatically run scripts through WSL.

For example, after reading `bd-generator`, call `run_skill_script` with `{"name":"bd-generator","file":"tools/flow_layout.py","args":["--recipe","navigate"]}`. The returned JSON contains drawing calls; the assistant must apply them with the Excel tools before the workbook changes.

For Python written by the assistant, use **`run_python`**. It writes `main.py` in a temporary working folder and runs Python directly on the user's machine with their permissions and installed libraries, including file and network access. Python must be installed locally (`python` on Windows, `python3` on Linux/WSL); `BRIDGE_PYTHON_BIN` can select its executable. Docker is not required. Ask/Auto controls execution approval.

Selected attachments are copied into the working folder (10 MB combined); requested output files return as `attachment://` references (10 MB combined). Execution is limited to 60 seconds and 512 KB of console output. Cancellation stops the process tree and the temporary working folder is removed after execution. These limits do not isolate Python from the rest of the computer. Example: `run_python {"code":"from pathlib import Path; Path('result.txt').write_text('Hello', encoding='utf-8')","outputFiles":["result.txt"]}`. Both Python runners use UTF-8, including subprocess output.

## AI capabilities and permissions

| Capability | Available route | Requirements / permission |
| --- | --- | --- |
| Read selections, cells, formulas, styles and objects | Excel MCP read tools | Connected workbook; some features require a recent Excel version |
| Write values/formulas; format, merge, resize, sort and filter | Excel MCP write tools | Ask/Auto controls supported changes and overwrite approvals |
| Manage worksheets, named ranges, validation and print layout | Excel MCP management tools | Ask/Auto; worksheet deletion requires approval in Ask mode |
| Create/edit tables, charts and pivot tables | Excel MCP object tools | Excel API support |
| Draw/edit shapes, connectors and flowcharts | Excel MCP shape tools | Excel API support; visual review for routed connectors |
| Run additional workbook operations | `execute_office_js` | Enabled by default; controlled by `BRIDGE_ALLOW_OFFICE_JS` / `allowOfficeJs`, plus Ask/Auto |
| Inspect attachments, closed Excel files and reference templates; import sheets | File/reference tools and `insert_sheets_from_file` | Selected files, configured references or access approval |
| Create/edit reusable instructions and run installed Python/Bash scripts | Skill tools and `run_skill_script` | Installed skill, local interpreter/dependencies; Ask/Auto for execution |
| Analyze data, generate files, access local files or call HTTP APIs | `run_python` | Local Python, libraries, OS/network access and service credentials when needed; Ask/Auto for execution |
| Search the web | Enable **Tools > Web search** beside the prompt | Uses Claude/Codex/Copilot native search; unavailable for OpenCode without an external search service |
| Automate a browser with Playwright | Enable **Tools > Playwright browser** beside the prompt | Local Node/npm and Chrome; first use may download Playwright MCP. Ask/Auto controls each browser action |

Choose **Ask** or **Auto** in the task pane to control the add-in's approval flow. Auto does not install software, grant administrator rights, add service credentials, enable provider-native tools or remove execution time limits. Python and skill scripts run with the bridge user's OS permissions; they can access files and the network through that execution route.

Removing the add-in prompt does not unlock provider-native tools. Claude and Copilot expose Excel MCP tools and optionally their native web tools; OpenCode allows Excel MCP tools and rejects other tool permission requests; Codex runs with a read-only sandbox and no approval escalation. Native web search is controlled separately by the Tools switch. Enabling native shell/browser tools or additional MCP servers requires changes to the relevant provider adapter and tool registration. Ask/Auto covers add-in tool calls, not arbitrary provider-native tool calls.

### Optional tools beside the prompt

Open **Tools** next to Ask/Auto and enable **Web search** and/or **Playwright browser**. Both default to off; choices are saved in this task pane and apply to the next message. Changing the selection reloads provider tool registration while preserving the native conversation where supported. Disabled tools are hidden from the model and direct calls are rejected by the bridge.

Web search uses only the selected AI provider's native tool: Claude Code's `WebSearch`/`WebFetch`, Codex's hosted search (`web_search = "live"`), or Copilot CLI's `web_search`/`web_fetch`. The add-in does not add a separate search endpoint or API key. Provider account/model availability and policies still apply. OpenCode's web search requires an external search backend, so this switch is disabled for OpenCode; no fallback service is configured.

Browser automation uses [Microsoft's Playwright MCP](https://playwright.dev/mcp/configuration/options), launched lazily as `npx --yes @playwright/mcp@latest --headless --isolated` on the bridge machine. Node/npm, network access for the first download, and Chrome must be available there. It uses a separate browser context for each connected workbook, without reusing your normal Chrome login. The assistant first lists available browser tools, then can navigate, inspect pages, click, fill forms and capture screenshots. **Ask** requests approval for each action; **Auto** approves it. Stop, tool/provider/chat changes, disconnect and bridge shutdown close the browser connection. Provider-native shell and other MCP servers retain their existing permissions.

If Chrome is installed in a nonstandard location, or you want to use an existing Chromium executable, set `BRIDGE_PLAYWRIGHT_EXECUTABLE` to its full path before starting the bridge. A missing browser is reported by the tool; enabling the switch alone does not install Chrome.

### Enable or disable skills

In **Settings > Manage skills**, use the **Enabled / Disabled** switch on any installed or packaged skill. The choice is saved on the bridge machine and applies across chats. Disabled skills stay installed and editable, but disappear from `/`, AI skill discovery and automatic template inventories. The bridge rejects loading their instructions, reading their templates or running their scripts until re-enabled. Instructions already present in chat history cannot be removed by this switch; start a new chat to avoid that old context.

## Commands

| Command | Description |
| --- | --- |
| `excel-partner setup` | Register the add-in, configure certificate trust and startup, and restart the bridge. |
| `excel-partner start` | Start the bridge after setup. |
| `excel-partner stop` | Stop the bridge. |
| `excel-partner update` | Install the latest npm release, refresh registration, and restart the bridge. |
| `excel-partner status` | Check whether the bridge is healthy. |
| `excel-partner uninstall` | Stop the bridge and remove add-in registration, startup entry, and certificate trust. |
| `excel-partner --help` | Show CLI usage. |

## Update

Save your work and close Excel before updating:

```powershell
excel-partner update
excel-partner status
```

Reopen Excel after updating. The command preserves your configuration, certificates, usage statistics, skills, and templates. It supports global npm installations; source installations use `scripts/update.ps1` or `scripts/update.sh`.

For an older release that does not yet have the `update` command, run `excel-partner stop`, `npm install -g excel-partner@latest`, then `excel-partner setup` once.

## Uninstall

Run the cleanup command before removing the npm package:

```powershell
excel-partner uninstall
npm uninstall -g excel-partner
```

Restart Excel. Local configuration, logs, skills, and templates remain on disk.

`execute_office_js` is enabled by default so skills can collect workbook data and run Office.js operations. Set `BRIDGE_ALLOW_OFFICE_JS=0` or `"allowOfficeJs": false` in `excel-partner.config.json` to disable it. Environment variables take precedence; restart the bridge and start a new chat after changing this setting.

## Session usage and export

Click the small info icon beside **Ask/Auto** to open session usage. **Session tokens** shows provider-reported token usage for the current chat and **AI time** shows cumulative time spent processing requests. The popover includes Input/Cache read/Output counts, uncached input and cache write when available, Copilot AI Credits, and **Export CSV**. Copilot accounting reads the exact CLI session’s `session.usage_record` entries in `session-state/<session-id>/events.jsonl` under `COPILOT_HOME` (default `~/.copilot`). It adds each usage record once and ignores duplicate shutdown/checkpoint totals. Original credit amounts are retained as integer `totalNanoAiu` (1 credit = 1,000,000,000 nano units), with up to nine decimal places in the display and CSV. No `/usage` command or extra model request is made. If the original log is absent, unreadable, or behind ACP, tokens fall back to ACP and credit/cache split data remains unavailable; it is never estimated. Old rounded credit records are explicitly labelled rounded, including reports mixing old and original data. The timer runs while a response is being generated, including provider startup, tools, and approval waits; idle time between requests is excluded. New chat starts a new native provider session and a separate accounting row. Reopening a saved chat or reconnecting resumes its native session ID; a failed resume is reported without silently creating a replacement. Provider/model changes resume the same native session where supported. Native session IDs are namespaced by provider, so different providers never share accounting.

**Export CSV** downloads one row per session with completed requests: session ID, title (when available in local history), timestamps, turn count, AI processing seconds, average seconds per turn, input/output/cache/total tokens, and token data completeness. Additional columns report known uncached input/cache write totals and coverage, original credit nano units and coverage, and Copilot CLI session IDs. Each new report row and History item identifies one native CLI session. Native titles are read from Copilot workspace metadata, Claude transcript title entries, Codex thread metadata, or OpenCode session export/list. If the provider has not assigned a title, the first user message remains a temporary title. Opening History refreshes available native titles without asking a model to generate one. Old entries without native IDs remain labelled legacy chats; their history and usage are preserved rather than guessed or reassigned. It includes requests that ended with an error or cancellation. The currently running request appears after it finishes. Old chats from before this feature have no retrospective usage data.

Input includes both cache read and cache write tokens; neither cache bucket is added again. With a complete split, total tokens = uncached input + cache read + cache write + output. Missing cache write/uncached values remain unknown, not zero. Copilot reads usage at turn completion, including failed/cancelled requests when records are available. Background title generation after a turn has finished is outside this report. Existing ledger records are preserved; original precision and missing CLI session IDs are not reconstructed for old records. The CSV identifies native CLI rows versus legacy chat rows. Missing token data is shown as `N/A`, or `partial` when only some turns have usage; CSV leaves unknown totals blank. Claude uses each turn’s original result usage: input includes uncached input plus cache read/write; output already includes reasoning. Codex uses native thread token updates, with a persisted rollout baseline before resume; reasoning is a subset of output and cache write is exposed when the CLI reports it. Copilot original log totals (or fallback ACP totals) are converted to per-turn increments. OpenCode reads original session export totals, where input excludes cache and output excludes reasoning: both cache buckets are added to input, and reasoning is added to output once. Its ACP prompt response only describes the last model call, so it is never used as a whole-turn total. Resumed cumulative counters are seeded before the next request; when a baseline is unavailable, that turn stays unreported rather than charging earlier history again. The popover and CSV expose original-data coverage and missing values.

Statistics are appended locally to `%LOCALAPPDATA%\Excel Partner\usage.jsonl` for the npm CLI, or `run/usage.jsonl` for a source checkout. Set `BRIDGE_USAGE_FILE` to override the path. The ledger contains IDs, provider/model, duration, result status, and token counts; chat history remains in the task pane's local storage. Deleting chat history does not delete the usage ledger.

## Local data and configuration

| Location | Contents |
| --- | --- |
| `%LOCALAPPDATA%\Excel Partner` | Bridge configuration, certificate, installed manifest, process ID, logs, and `usage.jsonl`. |
| `%USERPROFILE%\.excel-partner\skills` | User-installed skills. |
| `%USERPROFILE%\.excel-partner\templates` | Templates associated with skills. |

Optional settings for the npm installation belong in `%LOCALAPPDATA%\Excel Partner\config.json`. For example:

```json
{
  "referenceDirs": ["C:\\Templates"],
  "skillDirs": ["C:\\MySkills"]
}
```

Restart the bridge after editing configuration:

```powershell
excel-partner stop
excel-partner start
```

The bridge uses localhost HTTPS on port `8765`. Selected file contents and workbook context may be passed to the chosen AI provider as part of a request. Provider data policies apply. Review requested workbook changes and save important work before using AI editing tools.

## Troubleshooting

### PowerShell blocks npm scripts

If PowerShell reports that `npm.ps1` cannot run because script execution is disabled, use the Windows command shims:

```powershell
npm.cmd install -g excel-partner
excel-partner.cmd setup
```

The same `.cmd` suffix can be used for other npm-installed CLI commands.

### `excel-partner` is not recognized

Reopen PowerShell and check the global npm installation:

```powershell
npm.cmd list -g --depth=0
npm.cmd prefix -g
```

Confirm `excel-partner` is listed and that the directory returned by `npm prefix -g` is on your Windows `PATH`. If it was installed in WSL, install it again from Windows PowerShell.

### The Excel Partner button is missing

Run `excel-partner setup`, then save your work and fully exit all Excel windows before reopening Excel. Registration does not activate the add-in in an already running Excel process.

If the button is still missing, run the Microsoft sideload tool from PowerShell. It is downloaded on first use:

```powershell
npx.cmd --yes --package office-addin-dev-settings office-addin-dev-settings sideload "$env:LOCALAPPDATA\Excel Partner\manifest.xml" desktop --app Excel
```

If your organization blocks developer add-ins or certificate installation, ask your Microsoft 365 administrator about the applicable policy. This package registers a local sideloaded add-in.

### The task pane stays on Connecting or the bridge is unhealthy

```powershell
excel-partner status
excel-partner stop
excel-partner start
```

Open `https://localhost:8765/health` in a Windows browser and inspect the logs:

```powershell
Get-Content "$env:LOCALAPPDATA\Excel Partner\bridge.err.log" -Tail 50
Get-Content "$env:LOCALAPPDATA\Excel Partner\bridge.log" -Tail 50
```

If startup reports that port `8765` is already in use, stop the other Excel Partner bridge instance before retrying.

### The task pane shows a certificate error

Run `excel-partner setup` again to restore registration and certificate trust, then restart Excel.

If setup reports an expired certificate, a certificate that expires soon, or an incomplete certificate pair, save your workbook and close Excel, then recreate the certificate:

```powershell
excel-partner uninstall
excel-partner setup
```

### A provider shows Not installed, Login required, or a version error

Check the selected provider in Windows PowerShell:

```powershell
copilot --version
claude --version
codex --version
opencode --version
```

Only the provider you use needs to be installed. Install or update it using the command in the installation table, then sign in again. A CLI installed only in WSL cannot be used by the Windows bridge.

After installation or login, restart the bridge and select the provider again in Excel. Follow any minimum-version or account-access hint shown in the pane.

### Shared reference files

Open **Settings → Reference files → Add reference file**. In the popup, paste the full path of a local file (in Explorer: Shift + right-click → **Copy as path**), then enter a display name and description. References are shared across all skills and chats on this machine.

Only the name, path and description are stored in `~/.excel-partner/references.json`; the original file is never copied or uploaded. The assistant receives the catalog and reads relevant files on demand, using their latest contents. Registering a file grants read access to that exact file. Existing file-reading format and size limits still apply.

You can also give the assistant a local file path directly. `read_attached_file`, `read_attached_image`, and `read_workbook_file` request access when the file is outside configured workspace folders or saved references: **Ask** shows the file path for approval; **Auto** approves automatically. Approval applies to that exact file in the current connected session, without granting its folder or sibling files. Repeated reads reuse the grant. Denied requests do not grant access. Text/general file reads retain the 30 MB limit; supported images retain the 10 MB image limit.

Use **Edit** to change the description or choose a new location. Missing files are marked in the list. **Remove link** removes the saved metadata and its read access; it leaves the original file on disk.

### A file cannot be opened

Use **Choose local files** to attach it, or add its folder as a workspace. For persistent workspace access, add the folder to `referenceDirs` in the configuration file and restart the bridge.

An unsaved workbook or a workbook with a cloud URL may not have a local workspace folder. Use the file picker in that case.

### A write is refused with `would_overwrite`

The target cells already contain data. Confirm which cells you want replaced so the assistant can retry with overwrite permission.

### A tool reports `disconnected`, `timeout`, or `size_limit_exceeded`

Keep the task pane open and connected while a request is running. Close modal dialogs in Excel, select a smaller range, and retry after checking whether any changes were already applied. A disconnected write is not automatically replayed.

### The pane still shows an older version after updating

Run the update steps above, reopen the task pane, and start a new chat so the provider receives the current tool configuration.

## Development

From a source checkout:

```powershell
npm install
npm run build
npm run typecheck
npm test
```

The default test suite includes real Claude CLI tests against a fake API when Claude is installed. Real-provider model tests are opt-in through `LIVE_PROVIDERS`.

To build a distributable tarball:

```powershell
npm pack
```

`npm pack` builds the server and task pane through `prepack`. The package includes the compiled runtime and add-in assets; users do not need to build from source.
