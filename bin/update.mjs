import { spawnSync } from "node:child_process";
import { mkdirSync, readdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";

export const UPDATE_STATUS_FILE = "update-status.json";

/** Read the progress record of the last (or running) update; null when there is none or it is unreadable. */
export function readUpdateStatus(dataDir) {
  try { return JSON.parse(readFileSync(path.join(dataDir, UPDATE_STATUS_FILE), "utf8")); } catch { return null; }
}

/** Merge `patch` into the progress record. Written through a temp file so a reader never sees half of it. */
export function writeUpdateStatus(dataDir, patch) {
  mkdirSync(dataDir, { recursive: true });
  const next = { ...(readUpdateStatus(dataDir) ?? {}), ...patch, updatedAt: new Date().toISOString() };
  const file = path.join(dataDir, UPDATE_STATUS_FILE);
  writeFileSync(`${file}.tmp`, JSON.stringify(next, null, 2));
  renameSync(`${file}.tmp`, file);
  return next;
}

const CSHARP_BLOCK = "Add-Type -TypeDefinition @'\nusing System;\nusing System.Runtime.InteropServices;\nusing System.Text;\npublic static class EpCwd {\n  [DllImport(\"kernel32.dll\")] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);\n  [DllImport(\"kernel32.dll\")] static extern bool CloseHandle(IntPtr h);\n  [DllImport(\"kernel32.dll\")] static extern bool ReadProcessMemory(IntPtr h, IntPtr addr, byte[] buf, int size, out IntPtr read);\n  [DllImport(\"ntdll.dll\")] static extern int NtQueryInformationProcess(IntPtr h, int cls, byte[] info, int len, out int ret);\n  [DllImport(\"kernel32.dll\")] static extern bool IsWow64Process(IntPtr h, out bool wow);\n  // The working directory sits in the process's own memory (PEB -> ProcessParameters -> CurrentDirectory).\n  public static string Get(int pid) {\n    IntPtr h = OpenProcess(0x0410, false, pid);\n    if (h == IntPtr.Zero) return null;\n    try {\n      bool wow; IsWow64Process(h, out wow); if (wow) return null;\n      byte[] pbi = new byte[48]; int ret;\n      if (NtQueryInformationProcess(h, 0, pbi, pbi.Length, out ret) != 0) return null;\n      IntPtr peb = (IntPtr)BitConverter.ToInt64(pbi, 8);\n      byte[] b8 = new byte[8]; IntPtr n;\n      if (!ReadProcessMemory(h, IntPtr.Add(peb, 0x20), b8, 8, out n)) return null;\n      IntPtr parms = (IntPtr)BitConverter.ToInt64(b8, 0);\n      byte[] us = new byte[16];\n      if (!ReadProcessMemory(h, IntPtr.Add(parms, 0x38), us, 16, out n)) return null;\n      int len = BitConverter.ToUInt16(us, 0); IntPtr buf = (IntPtr)BitConverter.ToInt64(us, 8);\n      if (len <= 0 || len > 4096) return null;\n      byte[] str = new byte[len];\n      if (!ReadProcessMemory(h, buf, str, len, out n)) return null;\n      return Encoding.Unicode.GetString(str);\n    } finally { CloseHandle(h); }\n  }\n}\n'@";

/**
 * Stop every process that runs from, or sits in, the installed package: the bridge, the gateway and relay, helper
 * scripts, and the programs they started. Windows cannot replace a folder while any process has it as its working
 * directory or holds a file in it, and npm then fails with EBUSY. Programs the bridge starts inherit the package
 * folder as their working directory without naming it on their command line, and some outlive their parent, so three
 * things count: a command line naming the folder, a working directory inside it, and descendants of either.
 * An update or `npm install` of this package that is already running counts too: two of them in the same global folder
 * are what npm reports as a lock. The calling process, the `keep` pids and the calling process's own parents (the shell
 * that ran `update` can mention the package and `install` in its command line) are spared.
 */
export function stopPackageProcesses({ root, keep = [], packageName = "excel-partner", run = spawnSync, platform = process.platform }) {
  if (platform !== "win32") return { stopped: 0 };
  const spared = [process.pid, ...keep].filter(Number.isSafeInteger);
  const literal = (text) => `'${String(text).replace(/'/g, "''")}'`;
  const script = [
    CSHARP_BLOCK,
    `$root = ${literal(path.resolve(root))}`,
    `$pkg = ${literal(packageName)}`,
    `$spared = @(${spared.join(",")})`,
    "$self = $PID",
    "$all = @(Get-CimInstance Win32_Process)",
    "$ancestors = New-Object 'System.Collections.Generic.HashSet[int]'",
    `$cur = ${Number.isSafeInteger(process.pid) ? process.pid : 0}`,
    "while ($cur -gt 0 -and $ancestors.Add($cur)) { $row = $all | Where-Object { $_.ProcessId -eq $cur } | Select-Object -First 1; $cur = if ($row) { [int]$row.ParentProcessId } else { 0 } }",
    "$ids = New-Object 'System.Collections.Generic.HashSet[int]'",
    "function Spare($id) { $id -eq $self -or $spared -contains $id }",
    "foreach ($p in $all) {",
    "  $id = [int]$p.ProcessId; if (Spare $id) { continue }",
    "  $named = $p.CommandLine -and $p.CommandLine.IndexOf($root, [StringComparison]::OrdinalIgnoreCase) -ge 0",
    "  $cwd = [EpCwd]::Get($id)",
    "  $inside = $cwd -and $cwd.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)",
    "  $cl = [string]$p.CommandLine",
    String.raw`  $updater = $cl -match '[\\/](excel|office)-partner-update-[^\\/]+[\\/]update-runner\.mjs'`,
    String.raw`  $installer = ($cl -match 'npm-cli\.js"?\s+(install|i|add)\s') -and ($cl -match '(^|\s)(-g|--global)(\s|$)') -and $cl.IndexOf($pkg, [StringComparison]::OrdinalIgnoreCase) -ge 0`,
    "  $rival = (-not $ancestors.Contains($id)) -and ($updater -or $installer)",
    "  if ($named -or $inside -or $rival) { [void]$ids.Add($id) }",
    "}",
    "do { $before = $ids.Count; foreach ($p in $all) { $id = [int]$p.ProcessId; if (-not (Spare $id) -and $ids.Contains([int]$p.ParentProcessId)) { [void]$ids.Add($id) } } } while ($ids.Count -ne $before)",
    "foreach ($id in $ids) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }",
    "Start-Sleep -Milliseconds 800",
    "Write-Output $ids.Count",
  ].join("\n");
  // -EncodedCommand (UTF-16 base64) keeps the quoting of the paths and the C# block out of the way.
  const encoded = Buffer.from(script, "utf16le").toString("base64");
  const result = run("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded], { encoding: "utf8", windowsHide: true });
  const stopped = Number.parseInt(String(result.stdout ?? "").trim().split(/\s+/).pop() ?? "", 10);
  return { stopped: Number.isFinite(stopped) ? stopped : 0 };
}

/** What npm says when something has the folder open or a file is in the way. Anything else (offline, no such version) cannot be cured by trying again. */
const NPM_LOCK_ERROR = /\b(EBUSY|EPERM|EACCES|ENOTEMPTY|EEXIST|ELOCKED)\b|resource busy|locked|operation not permitted/i;
const INSTALL_ATTEMPTS = 3;
const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * npm renames the old package folder to `.<name>-<random>` while it replaces it and deletes it afterwards. A run that
 * failed leaves it behind, and it gets in the way of the next one.
 */
export function removeNpmLeftovers(globalRoot, packageName, fs = { readdirSync, rmSync }) {
  const pattern = new RegExp(`^\\.${packageName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}-[A-Za-z0-9]+$`);
  let removed = 0;
  try {
    for (const entry of fs.readdirSync(globalRoot)) {
      if (!pattern.test(entry)) continue;
      try { fs.rmSync(path.join(globalRoot, entry), { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }); removed++; }
      catch { /* still in use: npm will say so */ }
    }
  } catch { /* no global folder to look in */ }
  return removed;
}

/** Run setup from the newly installed package, after npm has replaced this one. */
export async function updatePackage({ root, dataDir, nodePath, stop, wasRunning, run = spawnSync, spec = "excel-partner@latest", report = () => {}, wait = pause }) {
  const packageName = "excel-partner";
  mkdirSync(dataDir, { recursive: true });
  // npm.cmd needs the Windows command shell. Every argument is a fixed literal, except `spec`, which only the CLI
  // (`update --from`) and the bridge's own configuration can set.
  const npmOptions = { cwd: dataDir, shell: true, windowsHide: true };
  const location = run("npm.cmd", ["root", "--global"], { ...npmOptions, encoding: "utf8" });
  if (location.error || location.status !== 0 || !location.stdout?.trim()) {
    throw new Error("Could not locate the global npm installation. Check npm is available on PATH.");
  }
  const installedRoot = path.resolve(location.stdout.trim(), "excel-partner");
  if (installedRoot.toLowerCase() !== path.resolve(root).toLowerCase()) {
    throw new Error("Update supports the global npm installation. For a source checkout, run scripts/update.ps1 or scripts/update.sh; for npx, rerun npx with excel-partner@latest.");
  }
  console.log(`Updating Excel Partner (${spec})…`);
  report("stopping");
  await stop();
  report("installing");
  // --force replaces files npm did not write (an older copy of a command it links) and --prefer-online skips a stale
  // answer to "what is the latest version"; both are what a forced update needs.
  const installArgs = ["install", "--global", "--force", "--prefer-online", "--no-audit", "--no-fund", /\s/.test(spec) ? `"${spec}"` : spec];
  let installed;
  for (let attempt = 1; attempt <= INSTALL_ATTEMPTS; attempt++) {
    if (attempt > 1) {
      console.log(`npm could not replace the package (attempt ${attempt - 1} of ${INSTALL_ATTEMPTS}). Closing whatever still holds it and trying again…`);
      await stop();
      await wait(attempt * 2000);
    }
    removeNpmLeftovers(path.dirname(installedRoot), packageName);
    installed = run("npm.cmd", installArgs, { ...npmOptions, stdio: ["inherit", "inherit", "pipe"], encoding: "utf8" });
    const errors = String(installed.stderr ?? "");
    if (errors) process.stderr.write(errors);
    if (installed.error || installed.status === 0 || !NPM_LOCK_ERROR.test(errors)) break;
  }
  const cli = path.join(installedRoot, "bin", "excel-partner.mjs");
  if (installed.error || installed.status !== 0) {
    report("recovering");
    if (wasRunning) {
      const recovery = run(nodePath, [cli, "start"], { cwd: dataDir, stdio: "inherit", windowsHide: true });
      if (recovery.error || recovery.status !== 0) console.error("Could not restart the bridge. After fixing npm, run excel-partner setup.");
    }
    throw new Error("Update failed during npm install. Check the npm error above. If it names EBUSY or EPERM, close Excel and any terminal or editor opened in the package folder, then run excel-partner update again.");
  }
  report("starting");
  const setup = run(nodePath, [cli, "setup"], { cwd: dataDir, stdio: "inherit", windowsHide: true });
  if (setup.error || setup.status !== 0) throw new Error("Package updated, but setup failed. Run excel-partner setup to finish registration and restart the bridge.");
  console.log("Update complete. Reopen Excel to load the updated add-in.");
  return { cli, installedRoot, version: installedVersion(installedRoot) };
}

/** Version of the package installed at `installedRoot`, or null. */
export function installedVersion(installedRoot) {
  try { return JSON.parse(readFileSync(path.join(installedRoot, "package.json"), "utf8")).version ?? null; } catch { return null; }
}
