import { spawnSync, type ChildProcess, type SpawnSyncReturns } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

/**
 * Provider CLIs this process started, also written to a file so that the next start can stop the ones a crash or hard kill
 * left running. A record names the process number, the time it started and its executable: Windows reuses process numbers,
 * so a number is only stopped when the process that has it now started at the recorded time with the recorded executable.
 */
export interface ProcessRecord { pid: number; startedAt: number; exe: string }

/** How far the start time Windows reports may be from the one recorded right after spawning. */
const START_TOLERANCE_MS = 5_000;

let registryFile: string | null = null;
const running = new Map<number, { child: ChildProcess; record: ProcessRecord }>();

/** Where the records are kept (the bridge's or gateway's data folder). Null keeps them in memory only, as in tests. */
export function setProcessRegistryFile(file: string | null): void {
  registryFile = file;
  persist();
}

function persist(): void {
  if (!registryFile) return;
  try {
    if (!running.size) { fs.rmSync(registryFile, { force: true }); return; }
    fs.mkdirSync(path.dirname(registryFile), { recursive: true });
    const temp = `${registryFile}.${process.pid}.tmp`;
    fs.writeFileSync(temp, JSON.stringify([...running.values()].map(entry => entry.record)), { mode: 0o600 });
    fs.renameSync(temp, registryFile);
  } catch { /* A record that cannot be written only loses the cleanup after a crash. */ }
}

const exeName = (file: string) => path.basename(file).toLowerCase().replace(/\.exe$/, "");

/** Record a provider CLI that was just spawned; the record goes away when it exits. */
export function trackProcess(child: ChildProcess): void {
  const pid = child.pid;
  if (!pid) return;
  running.set(pid, { child, record: { pid, startedAt: Date.now(), exe: path.basename(child.spawnfile ?? "") } });
  persist();
  child.once("exit", () => { if (running.get(pid)?.child === child) { running.delete(pid); persist(); } });
}

/** Stops a process and everything it started, now. Windows npm shims run the real CLI as a grandchild. */
export function killTreeSync(pid: number, run: typeof spawnSync = spawnSync): void {
  if (process.platform === "win32") run("taskkill", ["/PID", String(pid), "/T", "/F"], { windowsHide: true, stdio: "ignore", timeout: 5_000 });
  else { try { process.kill(pid, "SIGKILL"); } catch { /* already gone */ } }
}

/** Stop every tracked CLI before this process exits (the crash path; normal shutdown closes them one by one). */
export function killTrackedProcesses(): void {
  for (const { child, record } of running.values()) {
    if (child.exitCode === null && child.signalCode === null) killTreeSync(record.pid);
  }
  running.clear();
  persist();
}

type Observed = { pid: number; name: string; start: number };

/** What Windows reports now for these process numbers: executable name and start time. Empty when the lookup cannot run. */
function observe(pids: number[], run: typeof spawnSync): Observed[] {
  if (process.platform !== "win32" || !pids.length) return [];
  const filter = pids.map(pid => `ProcessId=${pid}`).join(" OR ");
  const script = `@(Get-CimInstance Win32_Process -Filter '${filter}' | ForEach-Object { [pscustomobject]@{ pid = $_.ProcessId; name = $_.Name; start = ([DateTimeOffset]$_.CreationDate).ToUnixTimeMilliseconds() } }) | ConvertTo-Json -Compress`;
  const result = run("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", Buffer.from(script, "utf16le").toString("base64")],
    { encoding: "utf8", windowsHide: true, timeout: 15_000 }) as SpawnSyncReturns<string>;
  if (result.error || result.status !== 0 || !String(result.stdout ?? "").trim()) return [];
  try {
    const parsed = JSON.parse(result.stdout) as Observed | Observed[];
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch { return []; }
}

/**
 * Stop the CLIs a previous run of this process recorded and never saw exit (it crashed or was killed), then forget the
 * records. Returns how many were stopped. Call once at startup, before any provider starts.
 */
export function reapRecordedProcesses(file: string, run: typeof spawnSync = spawnSync): number {
  let records: ProcessRecord[];
  try { records = JSON.parse(fs.readFileSync(file, "utf8")) as ProcessRecord[]; }
  catch { return 0; }
  const valid = Array.isArray(records) ? records.filter(record => Number.isSafeInteger(record?.pid) && record.pid > 0 && typeof record.exe === "string" && typeof record.startedAt === "number") : [];
  let stopped = 0;
  for (const live of observe(valid.map(record => record.pid), run)) {
    const record = valid.find(item => item.pid === live.pid);
    if (!record || exeName(live.name) !== exeName(record.exe) || Math.abs(live.start - record.startedAt) > START_TOLERANCE_MS) continue;
    killTreeSync(record.pid, run);
    stopped++;
  }
  try { fs.rmSync(file, { force: true }); } catch { /* the next start tries again */ }
  return stopped;
}
