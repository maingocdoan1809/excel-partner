# Installation and provider diagnostics

Run diagnostics before changing a CLI version or reinstalling the add-in:

These commands describe the Office Partner implementation on main-office.

```powershell
office-partner doctor
office-partner doctor gateway --json
office-partner doctor --provider copilot --probe --json
office-partner doctor addin --mcp --host excel
office-partner doctor --integrity --provider copilot --json
office-partner doctor --json --output partner-report.json
```

Office Partner builds also support `--host word` or `--host powerpoint` for those
Office MCP catalogs. Excel-only builds report a catalog host mismatch for these
hosts rather than claiming success. The MCP
check starts this product's local stdio server and lists its native tool schemas.
It does not call a tool, connect to a document, or check the official Claude
client's Office tools. Gateway Office execution remains in the official client.

The default command checks installed files, executable candidates, SDK metadata,
local certificate validity, Windows startup registration and the expected local
service identity. It does not launch provider CLIs. Missing providers are warnings
unless explicitly selected with `--provider` or enabled in gateway configuration.
Managed selections and activation failures are read from the runtime state;
the pane's current external provider selection is not available to the standalone doctor.

`--integrity` hashes all retained app inventories and managed CLI inventories
(filter CLIs with `--provider`). It detects dependency corruption as well as a
changed entrypoint, without launching providers. Large caches take longer to
scan. Missing management helpers or damaged metadata are reported explicitly.

`--probe` explicitly opts in to provider `--version` and existing sign-in status
commands. It can contact the provider's authentication service. It never starts
login, installs or updates a CLI, sends a model prompt, or performs inference.
Child environment variables disable provider self-update where supported; the
parent environment and the user's configuration files are not modified. Probe
support must still be validated against the CLI version in use.

Copilot has no reliable generic `auth status` command in this integration.
Doctor reports `AUTH_UNVERIFIED` instead of treating a config file as proof of a
valid sign-in. Other providers' successful sign-in checks do not establish model
access. Native provider handshake, model-driven tool execution, organization
policy and native resume require separate opt-in live validation.

When the external Copilot CLI differs from the SDK's bundled runtime metadata,
`SDK_CLI_PAIR_UNVERIFIED` records both versions. This is a compatibility warning,
not proof that the pair is broken. `CLI_VERSION_OK` means the CLI printed a
parseable version meeting the known adapter minimum; it does not mean the version
has passed a tool round trip. `MCP_CATALOG_OK` proves initialization and tools/list
only, not execution in Office or registration inside a provider runtime.

## When the bridge will not start

The separate entrypoint does not import the bridge, gateway bundles, provider SDKs
or certificate-generation dependency for its default checks:

```powershell
office-partner-doctor --json
node <package-directory>/bin/doctor.mjs --json
```

The `doctor.mjs`, `doctor-checks.mjs` and `doctor-probes.mjs` files must be present
together. MCP probes additionally need the installed MCP SDK and compiled MCP
entrypoint; Windows command shims in provider probes need cross-spawn. If these
are missing, diagnostics reports a failed check rather than loading the bridge.
Setup also copies a separate recovery launcher under application data; see below.

The core reads the product name from package.json and also supports the
`excel-partner` package on main without importing Office host implementations.
Neither command requires a pane to be open.

## Services that do not start after Windows restarts

The bridge and the gateway start when you sign in to Windows (Run entries that
call the recovery launcher with `start --at-logon`). The pane itself is served by
the bridge, so while the bridge is down the pane cannot show a button; use one of
these instead:

- **Start > Office Partner - Fix connection** (created by `office-partner setup`).
  It starts the bridge, and the gateway if it is set up. A bridge that is running
  but not answering is replaced. The window reports the result and stays open.
  The same repair is `office-partner fix`.
- `startup.log` in the application data folder has one line per start attempt
  made at sign-in or by Fix connection: time, service, attempt, outcome and a
  reason code (`NOT_ANSWERING`, `PORT_IN_USE`, `START_FAILED`, `NOT_SET_UP`,
  `RELEASE_DAMAGED`, `CERTIFICATE_PROBLEM`). A start at sign-in is tried up to
  three times, 15 seconds apart.
- `office-partner doctor` shows the latest outcome (`*.startup_last`) and checks
  that the Node.js and launcher files named by the sign-in entry still exist
  (`*.startup_target`; run `office-partner setup` again if one was moved).

A saved process number is never trusted on its own: Windows reuses numbers after a
restart, so a leftover record whose number now belongs to another program is
discarded and the service is started.

## JSON reports and exit codes

Reports contain `schemaVersion`, `diagnosticId`, timestamp, product/environment,
scope, overall status and checks. Each check has a stable `id`, `status`, `code`,
message, optional suggested action and allowlisted metadata.

| Exit | Meaning |
| --- | --- |
| 0 | No failed checks; warnings or checks not performed may remain. |
| 1 | At least one failed check. |
| 2 | Invalid arguments, output cannot be saved, or doctor itself cannot run. |

Statuses are `pass`, `warn`, `fail`, `skip`. Automation should inspect both exit
code and individual checks. A healthy HTTP endpoint must identify the expected
service, not just return HTTP 200. An unconfigured gateway reports `TLS_MISSING`;
use `doctor addin` when diagnosing an add-in-only installation.

`--output` saves the same JSON report to a new file. Existing files are never
overwritten. The file contains diagnostic metadata, not raw logs. Review it
before sharing. No upload is performed.

Reports omit raw CLI stdout/stderr, environment variables, credentials, account
identifiers, document/session names, prompts, tool inputs/results and personal
paths. Executable identity is represented by a hash, its source (PATH/override),
extension and candidate count. MCP metadata is limited to host, count and catalog
hash. Requests and process probes have time and output limits.

## Managed providers and migration

Interactive `office-partner setup` asks which providers to install. Enter keeps
existing CLIs. Noninteractive setup does not prompt or silently select a provider.
Choose explicitly with `--providers copilot,codex` or `--providers none`.

```powershell
office-partner providers list
office-partner providers install copilot
office-partner providers login copilot
office-partner providers remove copilot
office-partner versions
```

The release manifest pins Claude Code `2.1.294`, Copilot CLI `1.0.95` with SDK
`1.0.18`, Codex `0.162.1` and OpenCode `1.19.5`. These are explicit adapter
baseline pins, not a claim of production live certification. Installation checks
the package/version, full file inventory, exact SDK dependency and native adapter
initialization before activation. Preflight also checks sign-in (Copilot uses its
SDK's authenticated-status RPC rather than config-file presence). Native preflight never sends a model prompt.
It verifies the adapter's initialization path, which can differ by provider;
it does not prove model access or successful model-driven Office tool execution.

Full distributions are installed under `%LOCALAPPDATA%\Excel Partner\runtimes`.
Adapter code and dependencies live in the app release. Private CLIs
use absolute paths and child-only update suppression, with no global npm install
or PATH registration. Their configuration/history directories live separately
under `provider-data`, per provider version and platform. No tokens are copied
by Office Partner. A provider may require sign-in to its private CLI. Supported
OS credential stores or inherited provider authentication may remain shared.

Existing system CLIs are captured as legacy profiles when opting into managed
mode. A download, integrity, policy or login failure does not activate the staged
CLI. Sign in, then verify again; Copilot sign-in in Settings also verifies the
pending profile. Updating the system CLI does not select it for managed chats.
Legacy external profiles still depend on the user's external installation.

For the first migration from an older package, install this Office Partner build
and run `office-partner setup` while responses are finished. This establishes the
baseline snapshot and recovery startup entry. Older services without maintenance
controls cannot participate in a transactional switch; an update refuses that
switch and keeps the running copy. Run setup from the existing installation to
migrate it, then use the managed update flow for later releases.

New workers use the selected profile; running workers keep theirs. Native
conversation IDs are hashed in a profile sidecar so saved chats resume with the
same adapter/SDK/CLI and state directory. Older chats without a mapping use the
captured legacy CLI. Missing saved runtimes produce a repair error, never a
silent replacement chat. There is no automatic conversion of old provider history.

Fallback is limited to a new session's initial transport/protocol startup, or a
missing/modified active entrypoint. It selects another verified profile of the same provider.
Authentication, policy and model failures are reported. Once a prompt has been
sent, no runtime switch or automatic replay occurs.

The private CLI is isolated from ordinary user CLI updates, not protected from
the same Windows user's deliberate access. That user can locate and execute its
files. This is version isolation, not an OS security boundary.

Provider controls are documented by [GitHub](https://docs.github.com/en/copilot/reference/copilot-cli-reference/cli-command-reference),
[Claude Code](https://code.claude.com/docs/en/settings), and
[OpenCode](https://opencode.ai/docs/cli/). They supplement exact private installs;
they do not freeze provider services, models or organization policy.

## Update, going back and repair

```powershell
office-partner update
office-partner update --from office-partner@0.1.22
office-partner update --from C:\packages\office-partner-release.tgz
office-partner rollback 0.1.22
office-partner repair
office-partner debug on
office-partner debug off
office-partner diagnostics export partner-support.json
```

Updates prepare a complete app release with dependencies and selected provider
profiles alongside the running copy. Published dependencies are locked with
`npm-shrinkwrap.json`. Downloads and native preflight finish before draining
services. Authenticated service controls reject active turns, login attempts,
model probes and pending gateway tool results. Partial drain is undone on failure.
The gateway's official-client Office calls still execute in the official client;
none are redirected to the Office Partner pane.

The activation journal records the prior app/runtime selection. Startup failure
restores it, and a dead updater's interrupted switch is recovered on next start.
A damaged active release is not replaced from disk: no older release is kept.
Starting reports `RELEASE_DAMAGED`, and `repair` (or Start > Office Partner - Fix
connection) installs the same version again from npm.
If the original release is already damaged and a recovery target cannot start,
recovery reports the failure and keeps the verified target selected. It cannot
claim to have restarted a damaged copy successfully.

Only the running release is kept. The previous one stays on disk while an update
is switching, so a failed start or an interrupted update can be undone
automatically; once the new release is running (or the old one is restored) the
unneeded one and the runtimes made for it are deleted. Going back to an older
version is `office-partner rollback <version>`, which installs that version from
npm like `update --from office-partner@<version>`; it needs a network connection.
The `held` flag records a recovery. Updates always require an explicit user action.

Saved chats: a runtime profile made by a retired release is dropped when it
duplicates the current CLI (same version and data folder) and its chats move to the
current profile. A profile for an older CLI version that chats were created with is
kept, together with its CLI folder and data, and is served by the current release's
adapter; it goes away only when those chats no longer need it. Folders left behind by
an earlier cleanup that could not delete them are removed at the next update.

Repair validates the active release and managed CLI inventories, including
inactive profiles needed by old chats. It replaces damaged CLI distributions
while services are stopped; downloading replacements needs npm/network access.
A damaged app is installed again from npm. Missing backups or unrecoverable
historical adapter files require reinstalling that release; diagnostics does not
manufacture a compatible conversation runtime.

Settings > App & storage > System & recovery offers provider selection, system
and host-specific MCP checks, JSON export, repair and
debug logging. Destructive document actions are not part of these controls.
Debug expires after 15 minutes. Bridge event logs rotate at approximately 5 MiB
with three retained files. Raw local logs can contain provider error details;
diagnostic exports omit them and do not upload anything automatically.

## Recovery when the package cannot load

Setup retains a dependency-free management copy outside the npm package.
With Node.js still installed, run in PowerShell:

```powershell
$recovery = Join-Path $env:LOCALAPPDATA 'Excel Partner\recovery\recovery.mjs'
node $recovery versions
node $recovery repair
node $recovery fix
node $recovery start all
node $recovery diagnostics export partner-recovery.json
```

The stable launcher points to a complete immutable recovery toolset. Startup
registration uses this launcher. Metadata writes are atomic and retain a previous
copy; `repair` can restore valid metadata backups. A lost data directory or Node
installation still requires reinstallation. Use `office-partner uninstall all`
before `npm uninstall -g office-partner` to remove owned releases and CLIs. User
history, credentials, external CLIs and provider-data are retained deliberately.

## Validation boundary

Default automated checks use mock installers/providers and local MCP schema discovery.
They cover version isolation, pending login, dependency corruption, initial-only
fallback, saved native resume, busy services, failed activation, interrupted
updates, rollback by reinstalling and owner-aware uninstall. Separate opt-in Windows
checks installed real private CLIs and verified Copilot native registration,
real model-driven MCP execution in Chrome, continuation and History resume.
OpenCode's compatible pin was updated after live initialization checks. Private
Claude and Codex profiles without a sign-in remained pending as intended.
No Excel/Word/PowerPoint document write was performed. Before a broad rollout,
validate each signed-in profile on disposable Office documents. See the
[validation report](managed-runtime-validation.md) for exact coverage and limits.
The originally reported Copilot failure has not been reproduced from its error trace.

## Chat history storage

Chat history uses the bridge-owned SQLite database, with explicit pending-save errors and storage controls. See [durable history](history-storage.md) for migration, recovery and the supported Windows Node/driver pair.
