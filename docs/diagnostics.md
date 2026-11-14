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
office-partner providers remove copilot
office-partner versions
```

The release manifest pins Claude Code `2.1.294`, Copilot CLI `1.0.95` with SDK
`1.0.18`, Codex `0.162.1` and OpenCode `1.19.5`. These are explicit adapter
baseline pins, not a claim of production live certification. Installation checks
the package/version, full file inventory, exact SDK dependency and native adapter
initialization before activation. Preflight also checks sign-in (Copilot uses its
SDK's authenticated-status RPC rather than config-file presence); a provider that is not signed in
yet cannot be checked, which is not a reason to refuse its CLI. Native preflight never sends a model prompt.
It verifies the adapter's initialization path, which can differ by provider;
it does not prove model access or successful model-driven Office tool execution.

Full distributions are installed under `%LOCALAPPDATA%\Excel Partner\runtimes`.
Adapter code and dependencies live in the app release. Private CLIs
use absolute paths and child-only update suppression, with no global npm install
or PATH registration. Their configuration/history directories live separately
under `provider-data`, one folder per provider that is kept from one CLI version to the
next, so an update never signs anyone out. No tokens are copied
by Office Partner. A provider may require sign-in to its private CLI. Supported
OS credential stores or inherited provider authentication may remain shared.

A provider runs the CLI the product installed for it, or the system CLI when it installed none.
A release that ships the same CLI version reuses the install instead of downloading it again. A download,
integrity or policy failure leaves the CLI that was running in place and is recorded for `doctor`; it never
stops the update of the product itself. Updating the system CLI does not change a provider that runs a
private one.

For the first migration from an older package, install this Office Partner build
and run `office-partner setup` while responses are finished. This establishes the
baseline snapshot and recovery startup entry. Older services without maintenance
controls cannot participate in a transactional switch; an update refuses that
switch and keeps the running copy. Run setup from the existing installation to
migrate it, then use the managed update flow for later releases.

Every chat, new or saved, runs the CLI the provider is using now. A running session keeps the process it
started with until it ends. Removing a release on its own still requires reinstallation. Use `office-partner uninstall all`
before `npm uninstall -g office-partner` to remove owned releases and CLIs. User
history, credentials, external CLIs and provider-data are retained deliberately.

## Validation boundary

Default automated checks use mock installers/providers and local MCP schema discovery.
They cover reusing and switching CLI installs, a provider that is not signed in, dependency corruption,
busy services, failed activation, interrupted
updates, rollback by reinstalling and owner-aware uninstall. Separate opt-in Windows
checks installed real private CLIs and verified Copilot native registration,
real model-driven MCP execution in Chrome, continuation and History resume.
OpenCode's compatible pin was updated after live initialization checks. Private
Claude and Codex installs without a sign-in could not be checked further, as expected.
No Excel/Word/PowerPoint document write was performed. Before a broad rollout,
validate each signed-in profile on disposable Office documents. See the
[validation report](managed-runtime-validation.md) for exact coverage and limits.
The originally reported Copilot failure has not been reproduced from its error trace.

## Chat history storage

Chat history uses the bridge-owned SQLite database, with explicit pending-save errors and storage controls. See [durable history](history-storage.md) for migration, recovery and the supported Windows Node/driver pair.
