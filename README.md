# Excel Partner

Chat with an AI coding assistant, **Claude Code, GitHub Copilot CLI, Codex CLI or OpenCode**, to read, analyse and edit the workbook you have open in **Excel desktop on Windows**.

> Example: select a range and type *"Add a Total column in D. Calculate B times C for each row and bold the header."*

## How it works

```
Excel task pane  <-->  local bridge (localhost:8765)  <-->  your provider CLI (Claude / Copilot / Codex / OpenCode)
```

- The task pane is the chat UI inside Excel. Workbook changes run through Office.js in your open workbook.
- The bridge runs on your PC. It uses the provider CLI you already installed and signed in to. Excel Partner never asks for or stores an API key.
- Billing, usage limits and data policy belong to the provider you choose.

## Requirements

- Windows with Excel desktop (ExcelApi 1.9+; some features need a newer Excel)
- Node.js 20+ and npm, installed on Windows
- At least one provider CLI, installed and signed in on Windows (not only inside WSL)

## Quick start

Run in **Windows PowerShell**.

**1. Install a provider** (pick at least one):

| Provider | Install | Sign in |
| --- | --- | --- |
| Claude Code | `npm install -g @anthropic-ai/claude-code` | `claude` |
| GitHub Copilot | `npm install -g @github/copilot` | `copilot login` |
| Codex | `npm install -g @openai/codex` | `codex login` |
| OpenCode | `npm install -g opencode-ai` | `opencode auth login` |

**2. Install Excel Partner** (save your work and close Excel first):

```powershell
npm install -g excel-partner
excel-partner setup
excel-partner status
```

`setup` trusts a localhost certificate for your user, registers the add-in, starts the bridge and adds it to Windows start-up. A healthy bridge prints `Bridge running at https://localhost:8765`.

**3. Use it:** reopen Excel, open a workbook, choose **Home > Excel Partner**, pick a provider, select cells, type a request, press **Enter**. (**Ctrl+Enter** adds a new line.)

## What it can do

- **Read:** selections, ranges, formulas, formatting, tables, charts, pivots, shapes, and a rendered image of a range.
- **Write:** values and formulas (with a check for `#REF!`/`#NAME?` errors), formatting, sort, filter, merge, validation, conditional formatting, named ranges, print layout, sheets.
- **Build:** tables, charts, pivot tables, flowcharts with attached connectors.
- **Audit:** `find_errors` scans for formula errors; writes refuse to overwrite existing data unless overwrite is approved.
- **Extend:** your own reusable skills (`/skill-name`), reference files and templates, file attachments (`@`), optional web search and browser automation.
- **Work your way:** choose provider, model and effort per chat; queue a message while the assistant works; review changes and reopen past chats from History.

Type `/` in the chat to browse commands and skills of the selected provider.

## Testing from your own test-case file

Open any test-case workbook (your company's format, any folder) and let the assistant test it. Nothing is set up per project.

**One-time, per system you test (in a terminal):**

```powershell
excel-partner login admin https://staging.example.com   # Chrome opens; you sign in; press Enter in the terminal
excel-partner sessions                                   # list saved sign-ins (names, sites, date)
excel-partner logout admin                               # delete one
```

The password is typed into the site itself. Only the browser's session (cookies) is saved, in `%LOCALAPPDATA%\Excel Partner\sessions`, and the assistant never sees it. If the sign-in page is on another host (single sign-on), the hosts visited during login are saved with it; add more with `--origin https://sso.example.com`. A saved sign-in expires like any login: run the command again.

**Every time (in Excel):**

1. Turn on **Playwright** in Tools beside the prompt, then type `/qa` with what you want, for example `/qa test sheet "Test Cases" on https://staging.example.com`. Type `/qa login` to be shown the exact command for your system.
2. The assistant reads the sheet and tells you how it understood the layout (ID, steps, expected result, where results go). Correct it if it is wrong.
3. It opens the browser with your saved sign-in (`open_test_browser`). **You are asked every time, even in Auto mode**, and the dialog shows which sign-in and which sites. The browser can reach only those sites; everything else, IP addresses included, is blocked.
4. You watch every step in the live view above the message box. For each test case the assistant follows the steps, compares with the expected result and writes the outcome into the sheet (`record_test_results`): status, actual result, date and so on, in your own columns. Columns that do not exist yet, such as a progress column, are added after the last header with the same look.
5. At the end you get a summary: passed, failed, needs review, blocked.

**What the assistant remembers.** Everything it learns about a system (how to log in, where screens are, how to do a task, quirks, test data that is safe to keep) is saved as short Markdown notes in `%LOCALAPPDATA%\\Excel Partner\\knowledge\\<host>\\`, next to `sessions`. When the browser opens, the assistant is handed the list of notes for those sites (one line each, with the date) and opens only the ones it needs, so a new chat or a new day does not start from zero. In Ask mode you see each note before it is saved; notes that look like they contain a password, token or cookie are refused. A system keeps at most 30 notes (64,000 characters). Notes are treated as hints from earlier sessions, not as orders, and the assistant checks them against the page.

```powershell
excel-partner knowledge                      # which systems have notes (and where the folder is)
excel-partner knowledge https://staging.example.com           # list that system's notes
excel-partner knowledge show https://staging.example.com login # read one
excel-partner knowledge delete https://staging.example.com login
```

You can also edit the `.md` files by hand; the list is read from each file's header (`kind`, `summary`, `updated`).

Without a saved sign-in the assistant can still open a browser window and wait while you sign in by hand, but that sign-in is lost when the browser closes.

Cells that already hold something are not overwritten unless you agree; cells the assistant wrote earlier in the session (for example "Đang test" before "Pass") are updated freely. Results are written as text, never as formulas. A step that is unclear, or a page that behaves differently from the sheet, is marked for review instead of being guessed. If the same browser action fails twice, the assistant stops repeating it and reports.

## Safety and permissions

- **Ask / Auto** (next to the prompt): *Ask* approves each change, *Auto* approves automatically, including overwriting non-empty cells. In *Ask* mode, overwriting existing data needs your OK.
- Excel edits are **not transactional** and there is no undo for AI changes. Save important work first and review what was changed.
- Python and skill scripts (`run_python`, `run_skill_script`) run with **your Windows user permissions** and can reach local files and the network. Keep them on **Ask** unless you trust the task.
- `execute_office_js` (free-form Office.js) is on by default. Disable it with `BRIDGE_ALLOW_OFFICE_JS=0` or `"allowOfficeJs": false` in `config.json`, then restart the bridge.
- Selected cells and workbook context are sent to your chosen provider as part of each request. Their data policy applies.

### Instructions

**Settings > Instructions** adds your own rules for every chat (for example "answer in Vietnamese"). It is empty by default, so each provider keeps its native behaviour. **Use recommended** loads a ready-made set of Excel working rules (read before writing, use formulas, never overwrite silently, verify after editing). Press **Save**, then start a new chat so the old history does not keep earlier instructions. Note: with Claude, a non-empty instruction replaces its built-in prompt rather than adding to it.

## Commands

| Command | What it does |
| --- | --- |
| `excel-partner setup [addin\|gateway\|all]` | Register and start (default: add-in) |
| `excel-partner start` / `stop` / `restart` | Control the bridge |
| `excel-partner status` | Show what is running |
| `excel-partner doctor` | Check the installation and say what to fix |
| `excel-partner logs` | Show recent log lines |
| `excel-partner login <name> <url>` / `sessions` / `logout <name>` | Save, list and delete signed-in sessions for the test assistant (see Testing above) |
| `excel-partner knowledge [<url>]` / `show` / `delete` | See and manage what the assistant has learned about each system |
| `excel-partner update` | Install the latest release (close Excel first) |
| `excel-partner uninstall` | Stop, unregister and remove start-up and certificate trust |
| `excel-partner help` | Full CLI usage |

**Update:** `excel-partner update`, then reopen Excel and start a new chat. Configuration, skills and templates are kept.

**Uninstall:** `excel-partner uninstall`, then `npm uninstall -g excel-partner`, then restart Excel. Local data stays on disk (see below).

## Optional: Agent Gateway

The same package includes a local **Anthropic Messages-compatible API** backed by the provider CLIs you have installed (for example to sign in to Claude for Excel with a gateway). It does not need the add-in.

```powershell
excel-partner setup gateway     # prints the URL (https://localhost:8766) and a token
excel-partner info --copy       # show them again
excel-partner gateway models    # list advertised model IDs
```

Model IDs look like `claude/default`, `codex/default` or `provider/native-model`. Run `excel-partner doctor gateway` if something fails. `agent-gateway` is an alias for these commands.

## Data and configuration

| Location | Contents |
| --- | --- |
| `%LOCALAPPDATA%\Excel Partner` | Config (`config.json`), certificate, manifest, logs, `usage.jsonl` |
| `%USERPROFILE%\.excel-partner\skills` | Your skills |
| `%USERPROFILE%\.excel-partner\templates` | Templates used by skills |
| `%LOCALAPPDATA%\Agent Gateway` | Gateway config, certificate and token |

Optional `config.json` settings (restart the bridge after editing):

```json
{
  "referenceDirs": ["C:\\Templates"],
  "skillDirs": ["C:\\MySkills"]
}
```

Useful environment variables: `BRIDGE_ALLOW_OFFICE_JS`, `BRIDGE_PYTHON_BIN`, `BRIDGE_BASH_BIN`, `BRIDGE_PLAYWRIGHT_EXECUTABLE`, `BRIDGE_USAGE_FILE`.

The bridge listens only on `localhost:8765` over HTTPS. Chat history is stored in the task pane; the usage ledger (IDs, provider, model, duration, token counts) is local and contains no workbook content. Click the info icon beside **Ask/Auto** to see session tokens and export usage as CSV.

## Troubleshooting

Start with `excel-partner doctor`. Then:

| Problem | Fix |
| --- | --- |
| PowerShell says script execution is disabled | Use the shims: `npm.cmd install -g excel-partner`, `excel-partner.cmd setup` |
| `excel-partner` is not recognized | Reopen PowerShell. Check `npm.cmd prefix -g` is on `PATH`. Install from Windows PowerShell, not WSL |
| No **Excel Partner** button | Run `excel-partner setup`, then fully close **all** Excel windows and reopen. Still missing? `npx.cmd --yes --package office-addin-dev-settings office-addin-dev-settings sideload "$env:LOCALAPPDATA\Excel Partner\manifest.xml" desktop --app Excel`. Company policy may block developer add-ins |
| Task pane stuck on *Connecting* | `excel-partner stop`, `excel-partner start`, open `https://localhost:8765/health`, read `%LOCALAPPDATA%\Excel Partner\bridge.err.log`. If port 8765 is busy, stop the other instance |
| Certificate error | Run `excel-partner setup` again and restart Excel. If it says the certificate expired: `excel-partner uninstall`, then `excel-partner setup` |
| Provider shows *Not installed* / *Login required* | Check `claude --version`, `copilot --version`, `codex --version`, `opencode --version` in PowerShell, install or sign in again, restart the bridge and reselect the provider |
| `would_overwrite` | Target cells already hold data. Tell the assistant which cells it may replace |
| `disconnected`, `timeout`, `size_limit_exceeded` | Keep the pane open, close Excel dialogs, select a smaller range, and check whether part of the change was already applied. Writes are never replayed automatically |
| Old behaviour after updating | Reopen the task pane and start a new chat |
| File cannot be opened | Attach it with **Choose local files**, or add its folder to `referenceDirs` |

## License

MIT
