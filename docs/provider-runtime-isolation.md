# Provider runtime version isolation

Research date: 2026-11-10. Historical design notes from before implementation.
For the current Office Partner implementation and its validation limits, see
[runtime management and diagnostics](diagnostics.md).

## Recommendation

Give each Office Partner release a tested provider runtime profile. Resolve the
provider adapter, SDK, executable, launch arguments and protocol capabilities as
one unit. Use an application-owned installation with an exact version and an
absolute executable path. Keep installations for different versions side by side.

A user's global CLI update should then have no effect on the runtime used by
Office Partner. This isolates local executable changes; it cannot freeze remote
model behavior, service availability or organization policy.

## Findings in this repository

- `package.json` pins `@github/copilot-sdk` to `1.0.18`.
- `providers/src/copilot/adapter.ts` chooses `COPILOT_BIN` or the `copilot` PATH command.
  The ACP path uses this command for discovery and sessions.
- `providers/src/copilot/native.ts` explicitly passes the resolved external
  executable to `RuntimeConnection.forStdio` for models, commands and sessions.
  This overrides the SDK's bundled runtime selection.
- The installed SDK's package metadata associates SDK `1.0.18` with runtime
  `1.0.94`. Its TypeScript declarations allow `forStdio({ args })` without a path;
  its client then resolves the bundled runtime. Platform runtime packages are
  exact optional dependencies. The installed client adds `--no-auto-update`.
- `CopilotAdapter.MIN_VERSION` is `1.0.95`. The bundled `1.0.94` is therefore a
  candidate for investigation, not an already approved replacement. Changing
  only the native client's path would also leave detection and ACP on the global
  CLI, producing inconsistent results.
- `providers/src/base.ts` rejects versions below an adapter minimum but accepts
  newer versions without a tested compatibility profile. Its environment is
  `process.env`.
- Claude, Codex and OpenCode similarly use their `*_BIN` override or PATH.
- `bridge/src/provider-login.ts` independently chooses the Copilot executable.
- `shared/src/session-manager.ts` preserves native conversation IDs and
  reuses workers, but does not record a runtime profile/version in session state.

These facts establish exposure to global updates. They do not establish the
cause of the reported Copilot failure: the failing CLI version, transport, error
and tool registration trace have not been supplied or reproduced.

## Provider choices

| Provider | Proposed default | Work required |
| --- | --- | --- |
| Copilot | Runtime shipped with an exact SDK release, once validated | Select a tested SDK/runtime pair; align native sessions, detection and discovery. Check ACP separately because the SDK runtime must not be assumed to implement every CLI command. |
| Claude | Full versioned native distribution in an application-owned directory | Resolve the actual executable and required assets; disable updates for the child process; validate login and resume. |
| Codex | Full exact CLI distribution in an application-owned directory | Keep the current app-server adapter; validate its protocol and dynamic tools against the selected release. |
| OpenCode | Full exact CLI distribution in an application-owned directory | Select the matching ACP behavior, especially across major versions; verify the release's update and configuration behavior. |

Copilot's official local CLI documentation makes compatibility the application's
responsibility when supplying an external path. Its SDK offers a bundled runtime
as the default. See [GitHub's local CLI setup](https://github.com/github/copilot-sdk/blob/main/docs/setup/local-cli.md).

GitHub documents `COPILOT_AUTO_UPDATE=false`; this can supplement an isolated
installation but does not prevent another process replacing a global executable.
See [Copilot CLI reference](https://docs.github.com/en/copilot/reference/copilot-cli-reference/cli-command-reference).

Claude supports selecting a specific installer version and disabling background
updates. Its documentation distinguishes background updates from manual update
commands. Apply the supported controls to the managed child process and retain
an independent installation. See [Claude setup](https://code.claude.com/docs/en/setup).

OpenAI's deployment documentation recommends explicit approved CLI releases and
verifying the installed version for each update. An application-owned installation
can follow that principle while preserving this project's app-server transport.
See [OpenAI Windows deployment](https://learn.chatgpt.com/docs/enterprise/windows-deployment).

OpenCode documents upgrading to a specific version. Install selected versions in
separate directories instead of upgrading the active managed runtime in place.
See [OpenCode commands](https://opencode.ai/docs/cli/commands/).

## Shared resolver and profile

Introduce a shared resolver in `providers/src` used by both bridge and gateway.
Return an immutable descriptor containing provider ID, profile ID, adapter
revision, SDK version when applicable, exact CLI/runtime version, platform,
architecture, absolute entrypoint, arguments and child environment overrides.
Record the artifact source and integrity information in the release manifest.

An illustrative storage layout is:

```text
<application-data>/runtimes/<provider>/<exact-version>/<platform-arch>/
```

Install complete distributions, including loaders and native dependencies, into
staging directories. Verify artifacts and reported version before publishing
the directory. Installation must be serialized per artifact and recover cleanly
after interruption. Never overwrite a version used by an active worker. Retain
the previous approved profile; clean up only versions with no active users and
no retained conversation dependency. Handle Windows executable locks explicitly.

Every call path must use the same resolved profile: detection, login where
supported, models, commands, session creation, native resume and recovery. Where
a runtime lacks an interactive login command, use its documented authentication
flow and explicitly verify credential compatibility with the provider CLI.

Support managed mode as the product default and an explicit external executable
mode for users who need it. Preserve existing `*_BIN` configurations during
migration. In managed mode, do not silently fall back to PATH when an artifact
is missing. Provide installation/repair or report a precise error. Account for
the SDK's `COPILOT_CLI_PATH` override in a copied child environment so it cannot
silently redirect managed execution. Do not mutate the parent environment.

## Conversations, tools and credentials

Pin the runtime profile for the lifetime of a worker and persist its identity
alongside native conversation metadata. A new approved version applies to new
conversations; resuming an existing conversation selects its recorded profile.
Moving an existing conversation between versions requires verified provider
resume compatibility. Do not retry a failed prompt automatically after switching
versions: an earlier tool call may already have modified the document.

Executable isolation is only one layer. Shared provider state can still undergo
schema migrations when a global CLI updates. Investigate supported separation
of session/cache/config directories for each provider. Keep credentials under
provider control and use supported login/keychain mechanisms; do not copy tokens
into the runtime manifest or introduce API key storage. Verify authentication
and native resume with the chosen state layout before claiming full isolation.

Keep the Office MCP server owned by the bridge/runtime. External MCP tools
retain their real JSON Schemas and stable server namespaces; refresh native
registration when enabled tools change. Never replace them with generic wrappers.

Gateway Office tools remain supplied and executed by the official Claude Office
client. A shared runtime resolver must preserve this boundary for Excel, Word and
PowerPoint. A gateway upgrade must preserve pending tool results and conversation
continuations without replaying calls or binding to an Office Partner pane.

## Rollout and acceptance

1. Capture the failing Copilot CLI/SDK versions, chosen transport and exact error.
   Distinguish startup, policy, schema registration, tool execution and result
   conversion failures before selecting a replacement profile.
2. Validate a Copilot SDK/bundled-runtime pair with this adapter's required APIs.
   Align detection and discovery with that runtime, and retain ACP only with an
   explicitly validated executable/profile. Check packaged platform dependencies.
3. Add the shared resolver, versioned installer and persisted session profile.
   Migrate the other providers and existing executable overrides.
4. Gate promotion with mock protocol/contract checks: tool names and schemas,
   tool results, catalog refresh, cancellation, resume and gateway continuation.
   Cover missing artifacts, interrupted installs, environment overrides and
   active-worker retention. Run repository typecheck, build and tests.
5. Use opt-in live checks for authentication, model access and a tool round trip
   on each supported profile/host. A handshake or `--version` check alone cannot
   prove model-driven tool calling works. Use disposable documents for writes.
6. Demonstrate that changing PATH/global CLI versions leaves the managed profile
   unchanged, then demonstrate rollback for new sessions without replaying an
   existing turn. Report the active profile in diagnostics without secrets or
   personal paths.

This research used repository and installed SDK inspection plus official provider
documentation. No live CLI invocation, provider update, authentication change or
runtime compatibility test was performed.
