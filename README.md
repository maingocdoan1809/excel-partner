# Office Partner (formerly Excel Partner)

Chat with an AI coding assistant, **Claude Code, GitHub Copilot CLI, Codex CLI or OpenCode**, to read, analyse and edit the file you have open in **Excel, Word or PowerPoint desktop on Windows**. One install adds **Home > Office Partner** to all three.

> **Renamed.** The npm package is now `office-partner` and the command is `office-partner`. Already have `excel-partner`? Run `npm install -g office-partner` and `office-partner setup`, or just run `excel-partner` once (or press **Update now** in Settings): it installs `office-partner` and hands the command over. Your settings, certificate, backups and saved sign-ins are kept; the data folder is still `%LOCALAPPDATA%\Excel Partner`.

> Example: select a range and type *"Add a Total column in D. Calculate B times C for each row and bold the header."*

## How it works

```
Excel task pane  <-->  local bridge (localhost:8765)  <-->  your provider CLI (Claude / Copilot / Codex / OpenCode)
```

- The task pane is the chat UI inside Excel, Word or PowerPoint. Changes run through Office.js in the file you have open.
- The bridge runs on your PC. It can install private, pinned provider CLIs for each Office Partner release. Existing installations can keep using their system CLI during migration. Office Partner never asks for or stores an API key.
- Billing, usage limits and data policy belong to the provider you choose.

## Requirements

- Windows with Excel, Word and/or PowerPoint desktop (ExcelApi 1.9+, WordApi 1.1+, PowerPointApi 1.2+; many features need a newer version, and a tool that needs one says so instead of failing)
- Node.js 20.19+ (major 20), or 22.12+ below major 26, and npm, installed on Windows
- An account with at least one supported provider; sign-in stays under the provider CLI's control

## Quick start

Run in **Windows PowerShell**.

**1. Install Office Partner** (save your work and close Word, Excel and PowerPoint first):

```powershell
npm install -g office-partner
office-partner setup
office-partner status
```

`setup` trusts a localhost certificate for your user, registers the add-in, starts the bridge and adds it to Windows start-up. A healthy bridge prints `Bridge running at https://localhost:8765`.

**2. Choose providers.** Interactive setup asks which private CLIs to install (`copilot`, `claude`, `codex`, `opencode`). Enter keeps existing system CLIs. For an unattended selection use `office-partner setup --providers copilot,codex`; `--providers none` keeps external CLIs. System CLI files and PATH are never changed. Every provider signs in from the pane (Settings, Sign in): Copilot, Codex and Claude show a link and a code, and OpenCode lists its accounts (OAuth where it offers one, an API key otherwise). A private CLI that needs sign-in stays pending while the previous selection is retained.

**3. Use it:** reopen Excel, Word or PowerPoint, open a file, choose **Home > Office Partner**, pick a provider, select cells, text or a slide shape, type a request, press **Enter**. (**Ctrl+Enter** adds a new line.) Settings > App & storage > System & recovery also manages private providers. See [runtime management and recovery](docs/diagnostics.md).

## What it can do

- **Read:** selections, ranges, formulas, formatting, tables, charts, pivots, shapes, and a rendered image of a range.
- **Write:** values and formulas (with a check for `#REF!`/`#NAME?` errors), formatting, sort, filter, merge, validation, conditional formatting, named ranges, print layout, sheets.
- **Build:** tables, charts, pivot tables, flowcharts with attached connectors.
- **Audit:** `find_errors` scans for formula errors; writes refuse to overwrite existing data unless overwrite is approved.
- **Look across open files:** the button at the top left lists every file open with Office Partner, in any of the three apps. The assistant can read another one when you ask, for example to compare a workbook with a report or put a figure from Excel on a slide, or ask that file's assistant to make a change there. See *Connected files* below.
- **Read files that are not open:** Word documents and PowerPoint decks (including speaker notes) with `read_document_file`, Excel workbooks with `read_workbook_file`, from your reference folders or an attachment. Choosing a `.docx` or `.pptx` with the file picker or `@` works too.
- **Extend:** your own reusable skills (`/skill-name`), references (any file or folder) and templates, file attachments (`@`), optional web search and browser automation.
- **Work your way:** choose provider, model and effort per chat; queue a message while the assistant works; review changes and reopen past chats from History.

Type `/` in the chat to browse commands and skills of the selected provider.

### In Word

- **Read:** the outline, a section under a heading, search, comments and tracked changes.
- **Edit in place:** replace a passage (or all of it), lists, styles, tables, hyperlinks, bookmarks, footnotes, content controls, fields, tables of contents (desktop only), headers and footers, page layout, blank-line clean-up.
- **Review:** `propose_doc_edits` shows a card where you tick which edits to apply; nothing is written until you press Apply. Comments can be added, answered and resolved; track changes can be switched on, and accepted or rejected in bulk.

### In PowerPoint

- **Read:** the deck and its layouts, a slide's shapes and text, a rendered image of a slide.
- **Edit in place:** text, shapes, tables, order, duplicate, delete, background, alt text, links, z-order, grouping, alignment. `edit_slide_xml` changes what Office.js cannot reach by editing the slide's XML.
- **Check:** `verify_slides` looks for overlaps, shapes off the slide, overflowing and tiny text; `verify_slide_visual` shows the rendered slide.
- Not covered yet: speaker notes, pictures, charts, transitions, sections (use `execute_office_js` / `edit_slide_xml`).

There is no undo in Word and PowerPoint, so save the file first. Connected files and the testing tools work between Excel workbooks only.

## In a browser tab

Open `https://localhost:8765/` (or `/index.html`) in a browser for the same pane without Excel. It is the page Excel loads, served by the same bridge with your real providers, so chat, settings, sign-in, history and the assistant's reasoning all work. Reading or changing cells does not: there is no workbook, and the assistant is told so. A browser tab is not listed as a connected file, so other workbooks cannot send it requests. Like the add-in, it only answers pages from this computer.

## Connected files

Open two or more files, in any of Excel, Word and PowerPoint, and open **Home > Office Partner** in each. The button at the top left of the task pane (green dot = connected) lists them all, each with the app it is open in, and the current one marked **This file**.

Ask for it in plain words, for example *"Compare the totals in this sheet with the Actuals file"*, *"Put the revenue from Budget.xlsx on slide 3"* or *"In the Actuals file, fill column D with the totals"*. The assistant has two ways to reach the other file:

- **Look something up** (`list_connected_files`, `read_connected_file`): runs one **read** tool of that file's own app in the other file. It never writes. In *Ask* mode you are asked once per message and per file.

  | The other file is open in | Read tools |
  | --- | --- |
  | Excel | `get_sheets`, `read_range`, `get_cell_ranges`, `search_data`, `find_errors`, `check_formula_consistency`, `get_objects` |
  | Word | `get_doc_outline`, `read_doc_section`, `search_doc_text`, `read_comments`, `read_tracked_changes` |
  | PowerPoint | `get_presentation_metadata`, `list_slide_shapes`, `read_slide_text` |

- **Have something done** (`send_to_connected_file`): the assistant of this file sends a message to the assistant of the other file. That one runs it as a normal turn in **its own chat** (its own model, tools and Ask/Auto setting), and its final reply comes back to the asker. In *Ask* mode you see the exact message and are asked every time; *Auto* mode sends without asking.

What to expect when something is done in the other file:

- The request appears in the other file's chat as a message *From <app> · <file name>*, and its changes are listed in that file's **Changes** like any other edit. Save that file first.
- That file's own Ask/Auto setting applies. In *Ask* mode its change prompts appear **in that window**, so switch to it to approve them; the asker waits up to 10 minutes.
- If the other file fails or you stop it there, the asker is told it failed (with the reason), not that it finished. If the asker is stopped first, the request is not run.
- A file that is busy with its own request is refused at once; try again when it has finished. A file needs a provider selected in its own pane.

**Cell links (workbooks only) go exactly where they say, and nothing is guessed from the words around them.** The assistant points at cells with a markdown link that spells out the target: `[A1:E8](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8)` (`file` and `sheet` are optional; a space in a name is `%20`). Click it and that file's pane selects the range and its window comes to the front. The assistant does not have to work the target out: what it reads from or changes in another file comes back with ready-made links (`link`, `changed[].link`, taken from what the tools actually did) that it pastes into its reply. The same goes for this file: the result of a tool that works on a range (`read_range`, `write_range`, `get_selection`, …) carries a ready-made `link` for it. A cell written in plain words ("F10:F16", "Data!B2") still links to a sheet of the file you are in, but not in a reply that used the other-file tools, names another open file, or names one of this workbook's sheets ("sheet2/C5"): there it stays plain text, because a link to the wrong sheet or file is worse than no link. If a file is not open, or has no such sheet, you get a message saying so.
- A file's name is read from its app when someone needs it: when you open the list, or when the assistant lists or uses the files. After a Save As, the new name shows up the next time you open the list.
- Its task pane must be open: a file without the pane is not listed. A file that was closed and reopened gets a new id, so the assistant lists again.
- What is read or sent goes to the AI provider of the chat that handles it. Replies and cell contents from the other file are data, not instructions.

## Testing from your own test-case file

Open any test-case workbook (your company's format, any folder) and let the assistant test it. Nothing is set up per project.

**One-time, per system you test.** In Excel: Settings → *Testing: sign-ins, notes, scripts* → enter a name and the address → *Open browser to sign in* → sign in → *I'm signed in*. The same panel lists, shows and deletes saved sign-ins, notes and scripts, and opens their folders. Or in a terminal:

```powershell
office-partner login admin https://staging.example.com   # Chrome opens; you sign in; press Enter in the terminal
office-partner sessions                                   # list saved sign-ins (names, sites, date)
office-partner logout admin                               # delete one
```

The first sign-in downloads the browser driver (`playwright-core`) once into `%LOCALAPPDATA%\Excel Partner\tools`; it is not part of the add-in because most people never use it. The password is typed into the site itself. Only the browser's session (cookies) is saved, in `%LOCALAPPDATA%\Excel Partner\sessions`, and the assistant never sees it. If the sign-in page is on another host (single sign-on), the hosts visited during login are saved with it; add more with `--origin https://sso.example.com`. A saved sign-in expires like any login: run the command again.

**Every time (in Excel):**

1. Turn on **Playwright** in Tools beside the prompt, then type `/qa` with what you want, for example `/qa test sheet "Test Cases" on https://staging.example.com`. Type `/qa login` to be shown the exact command for your system.
2. The assistant reads the sheet and tells you how it understood the layout (ID, steps, expected result, where results go). Correct it if it is wrong.
3. It opens the browser with your saved sign-in (`open_test_browser`). **You are asked every time, even in Auto mode**, and the dialog shows which sign-in and which sites. The browser can reach only those sites; everything else, IP addresses included, is blocked.
4. You watch every step in the live view above the message box. For each test case the assistant follows the steps, compares with the expected result and writes the outcome into the sheet (`record_test_results`): status, actual result, date and so on, in your own columns. Columns that do not exist yet, such as a progress column, are added after the last header with the same look.
5. **Two things come out of every case.** The status in your sheet, and a Playwright script for that case. A script is saved only when the case passed, or failed and the assistant is sure the *application* is wrong (`save_test_script`). A case blocked by a technical problem, or one the assistant did not understand, is marked for review and gets no script.
6. Before it may say the run is done, the assistant must call `finish_test_run`. The bridge reads each row back from the sheet and refuses (with a list of what is missing) if a result is not there, a script is missing, or a script exists for a case that should not have one. At the end you get a summary: passed, failed, needs review, blocked.

**What the assistant remembers.** Everything it learns about a system (how to log in, where screens are, how to do a task, quirks, test data that is safe to keep) is saved as short Markdown notes in `%LOCALAPPDATA%\\Excel Partner\\knowledge\\<host>\\`, next to `sessions`. When the browser opens, the assistant is handed the list of notes for those sites (one line each, with the date) and opens only the ones it needs, so a new chat or a new day does not start from zero. In Ask mode you see each note before it is saved; notes that look like they contain a password, token or cookie are refused. A system keeps at most 30 notes (64,000 characters). Notes are treated as hints from earlier sessions, not as orders, and the assistant checks them against the page.

```powershell
office-partner knowledge                      # which systems have notes (and where the folder is)
office-partner knowledge https://staging.example.com           # list that system's notes
office-partner knowledge show https://staging.example.com login # read one
office-partner knowledge delete https://staging.example.com login
```

You can also edit the `.md` files by hand; the list is read from each file's header (`kind`, `summary`, `updated`).

**Regression scripts.** Each kept case becomes a file `%LOCALAPPDATA%\\Excel Partner\\scripts\\<host>\\<case id>.spec.js` (a Playwright Test file with a header: case, outcome, date, sheet row, the evidence). It asserts the *expected* result, so a script saved for an application defect fails until the defect is fixed. Scripts may import only `@playwright/test`, use relative paths and read only `BASE_URL` and `STORAGE_STATE`; ones with `require`, `fetch`, other sites or anything that looks like a password are refused. This is a check, not a sandbox: read a script before you run it. To run one (the add-in itself does not run them): `npm i -D @playwright/test`, then `BASE_URL=https://staging.example.com STORAGE_STATE=<saved session file> npx playwright test <file>`. `open_test_browser` also lists the scripts that already exist for the sites, so the assistant knows which cases were done before.

```powershell
office-partner scripts                                   # which systems have scripts (and where the folder is)
office-partner scripts https://staging.example.com        # list that system's scripts
office-partner scripts show https://staging.example.com TC-001
office-partner scripts delete https://staging.example.com TC-001
```

Without a saved sign-in the assistant can still open a browser window and wait while you sign in by hand, but that sign-in is lost when the browser closes.

Cells that already hold something are not overwritten unless you agree; cells the assistant wrote earlier in the session (for example "Đang test" before "Pass") are updated freely. Results are written as text, never as formulas. A step that is unclear, or a page that behaves differently from the sheet, is marked for review instead of being guessed. If the same browser action fails twice, the assistant stops repeating it and reports.

## Safety and permissions

- **Ask / Auto** (next to the prompt): *Ask* approves each change, *Auto* approves automatically, including overwriting non-empty cells. In *Ask* mode, overwriting existing data needs your OK.
- **No undo.** Excel's own `Ctrl+Z` cannot undo what the assistant does (Office.js writes do not go on Excel's undo stack), and Office Partner keeps no copy of your file, so large files stay fast. Save important work first and review what was changed in **Changes**.
- Large writes are sent in batches of about 15,000 cells with recalculation paused (the previous calculation mode is always restored), so Excel stays responsive.
- Python and skill scripts (`run_python`, `run_skill_script`) run with **your Windows user permissions** and can reach local files and the network. Keep them on **Ask** unless you trust the task.
- `execute_office_js` (free-form Office.js) is on by default. Disable it with `BRIDGE_ALLOW_OFFICE_JS=0` or `"allowOfficeJs": false` in `config.json`, then restart the bridge.
- Selected cells and workbook context are sent to your chosen provider as part of each request. Their data policy applies. Details: [docs/privacy-and-data-flow.md](docs/privacy-and-data-flow.md).
- Organizations: silent install, central manifest deployment and the admin policy file are described in [docs/enterprise-deployment.md](docs/enterprise-deployment.md) and [docs/admin-policy.md](docs/admin-policy.md).

### Instructions

**Settings > Instructions** adds your own rules for every chat (for example "answer in Vietnamese"). It is empty by default, so each provider keeps its native behaviour. **Use recommended** loads a ready-made set of working rules for the app you are in (in Excel: read before writing, use formulas, never overwrite silently, verify after editing). In Word and PowerPoint the add-in adds its own short rules for that app on top of yours (and uses them alone when yours is empty). Press **Save**, then start a new chat so the old history does not keep earlier instructions. Note: with Claude, a non-empty instruction replaces its built-in prompt rather than adding to it.

## Commands

| Command | What it does |
| --- | --- |
| `office-partner setup [addin\|gateway\|all]` | Register and start (default: add-in) |
| `office-partner start` / `stop` / `restart` | Control the bridge |
| `office-partner status` | Show what is running |
| `office-partner doctor` | Check the installation and say what to fix |
| `office-partner logs` | Show recent log lines |
| `office-partner login <name> <url>` / `sessions` / `logout <name>` | Save, list and delete signed-in sessions for the test assistant (see Testing above) |
| `office-partner knowledge [<url>]` / `show` / `delete` | See and manage what the assistant has learned about each system |
| `office-partner update [--from <spec\|.tgz>]` | Install the latest release (or the given package); same as Settings > Updates |
| `office-partner uninstall` | Stop, unregister and remove start-up and certificate trust |
| `office-partner help` | Full CLI usage |

**Update:** use **Settings > Updates** or `office-partner update`. A complete release, dependencies and selected provider versions are staged and verified before services stop. Busy services refuse the switch. If activation fails the error is reported and nothing goes back to the older release by itself; run the update again or `office-partner rollback <version>`. Once the new release is running, the old one is deleted: only the running release is kept on disk. Settings reports failures; the update log is `%LOCALAPPDATA%\Excel Partner\update.log`. Source checkouts can also stage managed releases. A global npm install updates the launcher; use `office-partner update` to switch the managed installation.

`office-partner update --from office-partner@<exact-version>` or `--from <path.tgz>` selects an artifact. `office-partner versions` shows the installed release and provider versions. Going back to an older version means installing it from npm: `office-partner rollback <version>` is the same as `office-partner update --from office-partner@<version>`. Saved conversations keep resuming with the CLI version they were created with while one is still in use; failed turns are never replayed during fallback. Settings > App & storage > System & recovery offers repair (a damaged installation is installed again from npm), diagnostics and temporary debug logging. If the pane does not open after a restart, use Start > Office Partner - Fix connection. See [diagnostics](docs/diagnostics.md) for migration, recovery commands and validation limits.

**Uninstall:** `office-partner uninstall all`, then `npm uninstall -g office-partner` (and `npm uninstall -g excel-partner` if the old name is still installed), then restart Word, Excel and PowerPoint. Managed app releases and CLIs are removed after both add-in and gateway are uninstalled. Uninstalling one retains shared runtimes for the other. User data, provider sign-ins/history, the recovery launcher and external CLIs stay on disk. Running only `npm uninstall -g` does not perform this application cleanup.

## External MCP servers

Settings → **MCP servers** lets you add, edit and test a local stdio command or a Streamable HTTP server. Enable a saved server in the composer's **Tools → Tools** popover for this connected file. Ask/Auto applies to external calls; external changes do not have Office undo. This first version does not support MCP authentication or OAuth.

The bridge serves Office Partner; the gateway is an API for the official Claude Office add-ins and preserves their client-owned Office tools. Shared MCP management does not require a document binding. See [MCP and gateway architecture](docs/mcp-runtime.md).

## Optional: Agent Gateway

The same package includes a local **Anthropic Messages-compatible API** backed by the provider CLIs you have installed (for example to sign in to Claude for Excel with a gateway). It does not need the add-in.

```powershell
office-partner setup gateway     # prints the URL (https://localhost:8766) and a token
office-partner info --copy       # show them again
office-partner gateway models    # list advertised model IDs
```

Model IDs look like `claude/default`, `codex/default` or `provider/native-model`. Run `office-partner doctor gateway` if something fails. `agent-gateway` is an alias for these commands.

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

The bridge listens only on `localhost:8765` over HTTPS. Chat history is stored in a local SQLite database managed by the bridge (with browser recovery copies for pending saves); the usage ledger (IDs, provider, model, duration, token counts) is local and contains no workbook content. Click the info icon beside **Ask/Auto** to see session tokens and export usage as CSV.

## Troubleshooting

Run `office-partner doctor --json --output partner-report.json` to save a diagnostic report, or `office-partner doctor --provider copilot --probe` to check that CLI's version and sign-in status without a model request. `office-partner doctor addin --mcp --host excel` checks the local MCP tool catalog. A standalone `office-partner-doctor` works when the bridge or its dependencies cannot load. See [diagnostics](docs/diagnostics.md) for scope, privacy and exit codes.

Start with `office-partner doctor`. Then:

| Problem | Fix |
| --- | --- |
| PowerShell says script execution is disabled | Use the shims: `npm.cmd install -g office-partner`, `office-partner.cmd setup` |
| `office-partner` is not recognized | Reopen PowerShell. Check `npm.cmd prefix -g` is on `PATH`. Install from Windows PowerShell, not WSL |
| No **Office Partner** button | Run `office-partner setup`, then fully close **all** Excel windows and reopen. Still missing? `npx.cmd --yes --package office-addin-dev-settings office-addin-dev-settings sideload "$env:LOCALAPPDATA\Excel Partner\manifest.xml" desktop --app Excel`. Company policy may block developer add-ins |
| Task pane stuck on *Connecting* | `office-partner stop`, `office-partner start`, open `https://localhost:8765/health`, read `%LOCALAPPDATA%\Excel Partner\bridge.err.log`. If port 8765 is busy, stop the other instance |
| Certificate error | Run `office-partner setup` again and restart Excel. If it says the certificate expired: `office-partner uninstall`, then `office-partner setup` |
| Provider shows *Not installed* / *Login required* | Check `claude --version`, `copilot --version`, `codex --version`, `opencode --version` in PowerShell, install or sign in again, restart the bridge and reselect the provider |
| `would_overwrite` | Target cells already hold data. Tell the assistant which cells it may replace |
| `disconnected`, `timeout`, `size_limit_exceeded` | Keep the pane open, close Excel dialogs, select a smaller range, and check whether part of the change was already applied. Writes are never replayed automatically |
| Old behaviour after updating | Reopen the task pane and start a new chat |
| File cannot be opened | Attach it with **Choose local files**, or add its folder to `referenceDirs` |

## License

MIT

See [durable history storage](docs/history-storage.md) for migration, pending-save recovery, export, backups and Windows Node compatibility.
