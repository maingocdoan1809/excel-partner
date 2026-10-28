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

## Test automation (QA)

Keep a test plan in Excel, let the assistant write Playwright scripts for it, run them against your **staging** system and get the results back in the workbook. Built for Odoo (form sign-in); other web apps work with `--preset generic`.

```powershell
excel-partner qa init C:qamy-project --url https://staging.example.com
cd C:qamy-project
excel-partner qa login admin           # you sign in once in a visible browser
excel-partner restart                  # the bridge picks up the project
```

Open `TestPlan.xlsx` in Excel, then in the chat use the skills installed by `qa init`:

| Skill | What it does |
| --- | --- |
| `/qa-testcases` | Turns the Requirements sheet into test cases (positive, negative, boundary, permission) with Req_ID links |
| `/qa-automate` | Looks at the real pages with your signed-in session, writes a script per test case, saves it as a **draft** and tries it while you watch |
| `/qa-heal` | Repairs scripts broken by UI changes or an unexpected extra dialog. Reports an application bug instead of weakening an assertion. `/qa-automate` and `/qa-report` run this repair themselves when a test fails (up to 2 rounds per test), so you do not have to call it |
| `/qa-report` | Runs the approved tests, writes Last_Run / Last_Result / Duration_s / evidence into the sheet by TC_ID, summarises failures |

RTM and Dashboard sheets in the template update themselves from the TestCases and Requirements sheets.

**You approve what runs.** A script is a draft until you read it and run `excel-partner qa approve TC_001`. Approval is tied to a hash of the script and the page objects it uses; any later change, by the assistant or by hand, withdraws it. Only approved scripts run unattended. The assistant has no tool to approve.

A repaired script becomes a draft again, so you read it and approve it once more. In Auto mode the assistant asks before each run of a draft; set `"run": { "healWithoutPrompt": true }` in `qa.config.json` to let it re-run a script it has just repaired (one that had been approved) without asking.

**Credentials never reach the assistant.** You type the password in the sign-in browser; only the resulting browser session is saved (`auth/`, git-ignored). `excel-partner qa secret set admin` can store the password encrypted for your Windows user (DPAPI) so expired sessions refresh by themselves. Passwords are never written to the workbook, the config or the chat.

**Guardrails.** Scripts are checked on every save and again before every run (no `fs`, `child_process`, `fetch`, `eval`, `page.route`, new browser contexts, cookies, absolute URLs or secrets in code). Runs use a scrubbed environment (no provider tokens), a fresh browser profile, and a browser that can reach **only** the origins in `qa.config.json`; everything else, IP addresses included, is refused. Screenshots are kept only for failures, traces only when asked, and results are pruned after 14 days. The checks are a safety net, not a sandbox: scripts still run with your Windows user permissions, so read a script before approving it. Approval does not protect against an assistant you have allowed to run its own Python in **Auto** mode; keep `run_python` on **Ask** for QA projects.

Commands: `qa init | use | doctor | upgrade | login | secret | status | list | validate | approve | revoke | run | clean | skills install | template`. Run `excel-partner qa help` for details, or see [docs/design/qa-harness.md](docs/design/qa-harness.md).

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
| `excel-partner qa <command>` | Test automation harness (see Test automation above) |
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
| `%LOCALAPPDATA%\Excel Partner\qa-secrets` | Stored QA passwords, encrypted with Windows DPAPI for your user (only if you ran `qa secret set`) |
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
