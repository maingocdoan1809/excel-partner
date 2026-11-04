import { spawnSync } from "node:child_process";
import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
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
 * The calling process and the `keep` pids are spared.
 */
export function stopPackageProcesses({ root, keep = [], run = spawnSync, platform = process.platform }) {
  if (platform !== "win32") return { stopped: 0 };
  const spared = [process.pid, ...keep].filter(Number.isSafeInteger);
  const literal = (text) => `'${String(text).replace(/'/g, "''")}'`;
  const script = [
    CSHARP_BLOCK,
    `$root = ${literal(path.resolve(root))}`,
    `$spared = @(${spared.join(",")})`,
    "$self = $PID",
    "$all = @(Get-CimInstance Win32_Process)",
    "$ids = New-Object 'System.Collections.Generic.HashSet[int]'",
    "function Spare($id) { $id -eq $self -or $spared -contains $id }",
    "foreach ($p in $all) {",
    "  $id = [int]$p.ProcessId; if (Spare $id) { continue }",
    "  $named = $p.CommandLine -and $p.CommandLine.IndexOf($root, [StringComparison]::OrdinalIgnoreCase) -ge 0",
    "  $cwd = [EpCwd]::Get($id)",
    "  $inside = $cwd -and $cwd.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)",
    "  if ($named -or $inside) { [void]$ids.Add($id) }",
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

/** Run setup from the newly installed package, after npm has replaced this one. */
export async function updatePackage({ root, dataDir, nodePath, stop, wasRunning, run = spawnSync, spec = "excel-partner@latest", report = () => {} }) {
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
  const installed = run("npm.cmd", ["install", "--global", /\s/.test(spec) ? `"${spec}"` : spec], { ...npmOptions, stdio: "inherit" });
  const cli = path.join(installedRoot, "bin", "excel-partner.mjs");
  if (installed.error || installed.status !== 0) {
    report("recovering");
    if (wasRunning) {
      const recovery = run(nodePath, [cli, "start"], { cwd: dataDir, stdio: "inherit", windowsHide: true });
      if (recovery.error || recovery.status !== 0) console.error("Could not restart the bridge. After fixing npm, run excel-partner setup.");
    }
    throw new Error("Update failed during npm install. Check the npm error above and retry excel-partner update.");
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
