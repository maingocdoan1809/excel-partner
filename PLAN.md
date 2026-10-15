# Excel Partner product plan

Excel Partner is a local Excel desktop add-in. The task pane sends workbook
operations through a loopback HTTPS/WebSocket bridge and a provider-neutral MCP
tool interface.

## Capabilities

- Read, edit, format and visualize the open workbook with Office.js.
- Use Claude Code, GitHub Copilot CLI or Codex CLI.
- Select models, attach user-selected files and control workbook permissions.
- Sign in to GitHub Copilot with a device code.
- Follow Excel's light/dark appearance, with manual theme overrides.
- Keep skills optional and user-installed; no project skills or company templates are bundled.

## Privacy and configuration

- Use neutral Excel Partner branding throughout the application and installers.
- Store local configuration in `excel-partner.config.json`.
- Read external folders only when the user explicitly grants access.
- Keep provider credentials under the provider CLI's control.
- Avoid shipping customer documents, company references, personal paths or account identifiers.

## Validation

Run `npm run typecheck`, `npm run build` and the relevant Vitest suites.
Use mock providers for automated checks. Live CLI checks require an installed
provider and the user's own authentication.
