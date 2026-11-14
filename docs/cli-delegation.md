# What the add-in leaves to the CLI

Goal: the add-in is an interface and a session manager. Whatever a CLI already does (model, effort, modes, permissions of its own tools, skills, web search, titles, usage), the CLI does; the add-in must not rebuild it. What stays in the add-in is what only the add-in knows: the Office tools and their guards, the chat history on screen, the host rules for Word and PowerPoint.

Status: **done** = in the code and tested; **checked** = read from the CLI on this machine; **assumed** = from the CLI's documented protocol, not run against the pinned version (needs a live check before it is switched on).

## Done: model and effort change inside the session

Before, any change of model or effort closed the CLI and opened it again (the chat survived through `resume`, but there was a pause, "Starting session…", and the CLI's cache was lost). Now each provider declares `liveModel` and `liveEffort` (`protocol/src/providers.ts`), and `ProviderSessionManager.ensure` restarts only for what a restart alone can change (instructions, web search, Playwright, another chat).

| Provider | Model | Effort | How |
|---|---|---|---|
| Copilot | live | live | SDK `session.setModel(model, { reasoningEffort })` (checked in the SDK types); ACP path: `session/set_config_option` (checked: the CLI offers `model`) |
| OpenCode | live | live | `session/set_config_option` for `model` and the effort option (checked) |
| Codex | restart | live | effort is already sent with every turn; model per turn is assumed (`turn/start` `model`) and not enabled until a signed-in check |
| Claude | restart | restart | both are launch flags; `set_model` as a control request is assumed |

## Checked on this machine: the CLIs describe their own options

`session/new` on Copilot and OpenCode returns `configOptions`, each with an id, a category and its values:

- Copilot: `mode` (agent, plan, autopilot), `model`, `allow_all` (permissions on/off).
- OpenCode: `model`, `mode` (build, plan); an effort option appears for models that have variants.

## Done: the CLI's own options are drawn in the pane

Each provider can report the settings its CLI offers for the open session (`ProviderAdapter.sessionOptions` / `setSessionOption`, one generic `provider_options` event, one route `POST /session/provider-option`). The pane shows them under the tools menu (the `+` beside the message box) and passes a choice back; it knows nothing about what an option means. A provider decides which categories it passes on (`providers/src/options.ts`): model and effort already have their own controls, and approval of Office edits is the add-in's.

- OpenCode: **mode** (build / plan), labelled by the CLI itself ("Session Mode"). Checked live: read, switched to plan and back, a value the CLI does not offer is refused.
- Copilot (ACP) reports `mode` (agent, plan, autopilot) and `allow_all`; the bridge uses Copilot's SDK, which has no equivalent call, so Copilot passes nothing yet. Codex (collaboration mode) and Claude (permission mode plan) can join by implementing the same two methods; neither could be run here (Codex is signed out, Claude is untested).
- A change applies to the open session. A restart (new instructions, web search, Playwright) opens the CLI at its default again, and the pane shows the new value at the next turn.

## Done: Claude keeps its own system prompt

The add-in's instructions are passed with `--append-system-prompt-file`, so Claude Code keeps its own tool guidance. The gateway, whose API clients send the whole prompt, still uses `--system-prompt-file` (replace). Not compared against real answers yet.

## Gateway

The gateway uses the same adapters and session manager, so it gets live model and effort (Copilot, OpenCode; effort on Codex) from its per-request route, and keeps replace mode for Claude. It ignores `provider_options` (an API client has no place to show them).

## Not done, and why

- **Skills.** They are a product feature, not a CLI setting: the add-in's skills carry project templates and scripts that run through the add-in's runner with an Ask prompt, because the CLIs' shell tools are off. The four CLIs look for skills in four different places and formats, so moving them is a rewrite of the feature, not a delegation. Codex already lists its own (`/skills`).
- **MCP servers.** The add-in's enabled list and namespaces are a stated requirement (AGENTS.md). Letting each CLI use its own MCP settings would drop that, so it needs an explicit decision.

## Stays in the add-in, on purpose

- **Ask/Auto for Office edits.** It guards the user's document (overwrite, delete, run script), which no CLI knows about. It is not a permission of a CLI tool. The CLI's own tool permissions (Copilot `allow_all`, OpenCode `permission`, Claude `--permission-mode`) are a different thing and could be shown through item 1.
- **The locked-down launch** (Claude `--tools`, `--setting-sources project`, `--strict-mcp-config`; Copilot `--no-custom-instructions`, `--available-tools`; Codex `sandbox: read-only`, `approvalPolicy: never`). It keeps a CLI embedded in Office from touching files or reading the user's own CLI configuration. Relaxing it (for example reading the user's `CLAUDE.md`/`AGENTS.md`) is a product decision.
- **The history database, the selection snapshot, the host rules, Playwright** (no CLI has them).

## Already left to the CLI

Model catalog (`listModels`), slash commands (native discovery), chat titles (native), web search (the CLI's own tool; the toggle only switches it on), resume, sign-in and sign-out, usage where the CLI reports it.
