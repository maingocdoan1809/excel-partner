# Installing Office Partner (Excel Partner)

Target environment: **Word, Excel and PowerPoint desktop on Windows** + **WSL2**
running the bridge and the provider CLIs. One install covers all three apps.
Office on macOS and multi-user scenarios are not supported yet; the web
versions of Word and PowerPoint are untested (see "Known limits" below).

## 1. Prerequisites

- Windows 10/11 with Word, Excel and/or PowerPoint desktop (Microsoft 365 or 2016+; install only the apps you use).
- WSL2 with a Linux distro, Node.js 20+.
- At least one provider CLI installed and logged in **inside WSL**:
  - Claude Code: `npm install -g @anthropic-ai/claude-code`, then `claude`
  - GitHub Copilot: `npm install -g @github/copilot`, then `copilot`
  - Codex: `npm install -g @openai/codex`, then `codex login`
  - OpenCode: `npm install -g opencode-ai`, then `opencode auth login`

The add-in reuses each CLI's existing login. It never asks for or stores an API
key.

## 2. Build the bridge (WSL2)

```bash
cd /home/<you>/excel-ai-addin
bash scripts/setup.sh
```

This installs dependencies, generates `certs/localhost.crt(.key)` and builds
`dist/` (bridge + MCP server) and `addin/dist/taskpane.js`.

## 3. Trust the localhost certificate

The task pane and API share one HTTPS origin (`https://localhost:8765`).

- If `mkcert` was available, its root CA is already trusted on Windows.
- Otherwise, copy `certs/localhost.crt` to Windows and import it:

```powershell
certutil -addstore -user Root <path-to>\localhost.crt
```

Confirm `https://localhost:8765/health` opens in a Windows browser without a
certificate warning.

## 4. Start the bridge

```bash
bash scripts/start.sh
curl -k https://localhost:8765/health
```

The bridge listens on `127.0.0.1:8765` by default. Windows reaches WSL
`localhost` through WSL's port forwarding; no extra firewall rule is normally
needed.

## 5. Sideload the add-in

From Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\sideload.ps1 `
  -ManifestPath '\\wsl.localhost\<Distro>\home\<you>\excel-ai-addin\addin\manifest.xml'
```

Or, from WSL, do steps 3-5 in one go: `bash scripts/install-windows.sh --open`.

Then in Word, Excel or PowerPoint (restart any window that was already open): **Home > Office Partner**.
If the ribbon button is missing, use **Insert > My Add-ins > Developer Add-ins > Office Partner**.
The registration is shared by all Office apps, so there is nothing to repeat per app.

## 6. First run

1. The task pane opens and shows `connected` in the top-right.
2. Pick a provider. The pane shows the detected version and login status.
3. Select a range (Excel), some text (Word) or a slide/shape (PowerPoint), type a request and press **Send** (or Ctrl+Enter).

## Upgrade / uninstall

```bash
bash scripts/update.sh     # pull, reinstall, rebuild, restart
bash scripts/uninstall.sh  # stop + remove build output and certs
```

Remove the add-in from Excel via **Insert > My Add-ins** and delete the trusted
catalog folder if you want a full cleanup.

## Known limits for Word and PowerPoint

- **Undo.** Excel keeps an undo journal. Word and PowerPoint only get a backup copy (`.docx` / `.pptx`) before the
  first change of each reply; there is no automatic restore. List the copies with `office-partner backups`.
- **Office version.** Each tool checks the API level it needs and answers `unsupported` with a hint when the
  Office build is too old (for example PowerPoint slide export needs PowerPointApi 1.8, Word tracked changes need
  WordApi 1.6). Table of contents and page layout in Word need the desktop app (WordApiDesktop).
- **Not done yet (phase 2).** PowerPoint speaker notes, inserting images or icons, charts, slide master edits,
  transitions and sections.
- **Web.** Backups use `getFileAsync(Compressed)`, which is unverified in Word and PowerPoint on the web.
- **Connected files** work between Word, Excel and PowerPoint: the assistant can read another open file (read-only, with your approval) or ask that file's assistant to change it. Cell links work for workbooks only.
- **Files that are not open**: `read_document_file` reads Word and PowerPoint files (speaker notes included) and `read_workbook_file` reads Excel files, from your reference folders or an attachment.
