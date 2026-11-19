# Enterprise deployment

Office Partner runs per Windows user:

- a local bridge on `https://localhost:8765`;
- a self-signed localhost certificate trusted in that user's `CurrentUser\Root` store;
- an add-in manifest that points to the bridge.

No administrator rights are needed, and nothing is installed machine-wide except the optional admin policy file.

## Supported

- Windows 10 and 11, with Word, Excel and PowerPoint desktop (Microsoft 365 Apps or Office 2021 and later).
- Not supported: Office on the web, Office for Mac.
- Node.js 22.12 or later, before 26. The installer adds a portable Node.js for the user when none is found.
- One signed-in user at a time on a machine. The bridge and gateway use fixed ports (8765 and 8766).
  - On Remote Desktop servers and with fast user switching, the bridge refuses connections from other users' processes.
  - Only the first user can run it at a time.

## MSI installer

A per-user MSI (about 8 MB) is built by `.github/workflows/msi.yml` (`npm run build:msi` locally, with the WiX 5.0.2 CLI) for x64 and ARM64. It is the same install as the script below, packaged for Intune and SCCM:

```bat
msiexec /i office-partner-0.1.26-x64.msi /qn /l*v install.log PROVIDERS=copilot NOSIDELOAD=1
```

- Run it in the user's context. `PROVIDERS` defaults to `copilot`; `NOSIDELOAD=1` is for a manifest deployed from the Microsoft 365 admin center.
- Inside the MSI: the add-in, its dependencies pinned by `npm-shrinkwrap.json`, and the Windows SQLite and koffi binaries. A computer therefore needs no GitHub download and no compiler.
- Downloaded while installing, pinned in `bin/runtime-manifest.json` with a SHA-256 for the Node.js zip: Node.js (into `%LOCALAPPDATA%\Programs\nodejs`) and the Copilot CLI. Behind a firewall, allow nodejs.org and registry.npmjs.org; `OP_NODE_MIRROR` and npm's `registry` setting point to internal mirrors.
- It always uses its own pinned Node.js, never one already on the computer, because the bundled SQLite binary is built for that Node.js release.
- It installs into `%APPDATA%\npm\node_modules\office-partner`, so `office-partner update` and `rollback` keep working. Updates change the files without changing the MSI's version in Apps & features.
- Not included: Python (`run_python`, skills with scripts), Git Bash, Chromium, and the Claude, Codex and OpenCode providers (`office-partner providers install <provider>` adds one). Those features say what is missing when used.
- The install log is `%LOCALAPPDATA%\Excel Partner\msi-install.log`.

### Uninstall and erase local data

Close Word, Excel and PowerPoint. Run this as the user who installed Office Partner, using the installed product's MSI or ProductCode:

```bat
msiexec /x "office-partner-<version>-x64.msi" PURGEDATA=1 /l*v uninstall.log
```

`PURGEDATA=1` permanently deletes `%LOCALAPPDATA%\Excel Partner` (chats, settings, logs and managed CLIs) and `%USERPROFILE%\.excel-partner` (including personal skills and templates). It also deletes `%LOCALAPPDATA%\Agent Gateway` if the MSI installed the Gateway feature. Back up any files you want to keep before running the command. A purge failure is logged to `%TEMP%\office-partner-msi-purge.log`.

Uninstalling from Windows Settings preserves these folders. MSI upgrades also preserve them. The purge covers the listed default folders. Data relocated through environment overrides and credentials held by separately installed provider CLIs can be removed separately.

## Silent install (Intune, SCCM, scripts)

Run `install-office-partner.cmd` **in the user's context** (Intune: *Install behavior = User*). Run as SYSTEM, it exits with code 60.

```bat
install-office-partner.cmd /quiet /version 0.1.23
```

| Switch or variable | Meaning |
|---|---|
| `/quiet` or `OP_QUIET=1` | No prompts and no pause. Exits with 50 if Word, Excel or PowerPoint is open. |
| `/version x.y.z` or `OP_VERSION` | Release to install. Without it, the latest release is installed (not recommended for managed rollouts). |
| `OP_PROVIDERS` | Provider CLIs to install, comma separated (default `copilot`). |
| `OP_NODE_VERSION` | Exact portable Node.js version, verified against the SHA256 list from nodejs.org. |
| `OP_REGISTRY` | Internal npm registry. |
| `OP_NO_SIDELOAD=1` | Do not register the add-in per user (the manifest comes from Microsoft 365, see below). |
| `/uninstall` | Remove the add-in registration, startup entry, certificate trust, package and the portable Node.js the installer added. |
| `OP_PURGE=1` | With `/uninstall`: also delete chats, logs and provider sign-ins. |

Running the installer again with a newer `/version` switches the active release, the same as `office-partner update`.

| Exit code | Meaning |
|---|---|
| 0 | Success |
| 10 | Node.js download or checksum failed |
| 20 | `npm install` failed (network, proxy, registry) |
| 30 | `office-partner setup` or `update` failed |
| 40 | Windows did not trust the localhost certificate |
| 50 | Office is running (quiet mode) |
| 60 | Not run as a signed-in user (SYSTEM) |
| 1 | Anything else |

The install log is `%LOCALAPPDATA%\Excel Partner\install.log`.

## Deploying the manifest centrally

Developer sideloading (the per-user registry key used by default) can be blocked by policy. Instead:

1. Upload `frontend/manifest.xml` from the installed package in the Microsoft 365 admin center: *Settings > Integrated apps > Upload custom apps*.
2. Assign it to the users.
3. Install with `OP_NO_SIDELOAD=1` (or run `office-partner setup --no-sideload`). Each machine still needs the per-user bridge and certificate.

## Certificate

- Each user gets their own certificate: a leaf, not a CA, valid for `localhost`, `127.0.0.1` and `::1`, for 365 days.
- It is renewed automatically 60 days before it expires, at the next start (which runs at sign-in). Windows asks the user once to trust the new certificate. The old certificate's trust is then removed.
- If a policy blocks user root certificates, setup fails with exit code 40.

## Proxy and internal registry

- Update lookups and installs use npm itself, so npm's `proxy`, `https-proxy`, `registry` and `cafile` settings apply. Set them with `npm config set`, or with the policy key `registry`.
- The SQLite driver (`better-sqlite3`) downloads a prebuilt binary from GitHub when it is installed. Behind a firewall, set `npm_config_better_sqlite3_binary_host` to an internal mirror, or allow GitHub release downloads.
- Without the driver, chat history is off and everything else works.

## Updates and rollback

- Updates install a whole release next to the current one and switch only when it is complete.
- `office-partner update` refuses:
  - an older version (unless `--allow-downgrade`, or `rollback <version>`);
  - a deprecated version.
- The admin policy can turn updates off or pin a version for the fleet (see [admin-policy.md](admin-policy.md)).
- To withdraw a bad release:
  1. Deprecate it on npm, so machines that have not updated refuse it.
  2. Pin the previous version in the policy, so machines that already updated go back at their next update.

## Restricting features

See [admin-policy.md](admin-policy.md). What leaves the machine is described in [privacy-and-data-flow.md](privacy-and-data-flow.md).
