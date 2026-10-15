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
- Choose a provider, model, and reasoning effort from the task pane.
- Attach local files or browse workspace files with `@`.
- Add reusable instructions through **Settings > Manage skills**, then invoke them with `/skill-name`.
- Review workbook changes and return to conversations through chat history.

## Commands

| Command | Description |
| --- | --- |
| `excel-partner setup` | Register the add-in, configure certificate trust and startup, and restart the bridge. |
| `excel-partner start` | Start the bridge after setup. |
| `excel-partner stop` | Stop the bridge. |
| `excel-partner status` | Check whether the bridge is healthy. |
| `excel-partner uninstall` | Stop the bridge and remove add-in registration, startup entry, and certificate trust. |
| `excel-partner --help` | Show CLI usage. |

## Update

Save your work and close Excel before updating:

```powershell
excel-partner stop
npm install -g excel-partner@latest
excel-partner setup
excel-partner status
```

Reopen Excel after setup. Running `setup` again refreshes the registered manifest and startup path.

## Uninstall

Run the cleanup command before removing the npm package:

```powershell
excel-partner uninstall
npm uninstall -g excel-partner
```

Restart Excel. Local configuration, logs, skills, and templates remain on disk.

## Local data and configuration

| Location | Contents |
| --- | --- |
| `%LOCALAPPDATA%\Excel Partner` | Bridge configuration, certificate, installed manifest, process ID, and logs. |
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
```

Only the provider you use needs to be installed. Install or update it using the command in the installation table, then sign in again. A CLI installed only in WSL cannot be used by the Windows bridge.

After installation or login, restart the bridge and select the provider again in Excel. Follow any minimum-version or account-access hint shown in the pane.

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
