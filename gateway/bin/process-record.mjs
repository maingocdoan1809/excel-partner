// Decides what to do with a service's leftover "process record" (a saved process number).
// Shared by the add-in CLI and the gateway CLI; needs only Node built-ins.
import { spawnSync } from "node:child_process";

/**
 * Who a recorded process number belongs to now. Windows hands the same numbers out again after a restart, so a number
 * that is "alive" may be an unrelated program. Returns "ours" (Node, with a command line the caller recognises),
 * "other", or "unknown" when the lookup itself could not run.
 */
export function processOwner(pid, isOurCommandLine, { run = spawnSync, platform = process.platform } = {}) {
  if (platform !== "win32" || !Number.isSafeInteger(pid) || pid <= 0) return "unknown";
  const script = `Get-CimInstance Win32_Process -Filter 'ProcessId = ${pid}' | Select-Object Name,CommandLine | ConvertTo-Json -Compress`;
  const observed = run("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", Buffer.from(script, "utf16le").toString("base64")],
    { encoding: "utf8", windowsHide: true, timeout: 8000 });
  if (observed.error || observed.status !== 0) return "unknown";
  if (!String(observed.stdout ?? "").trim()) return "other";
  let info; try { info = JSON.parse(observed.stdout); } catch { return "unknown"; }
  return /^node(?:\.exe)?$/i.test(info?.Name ?? "") && isOurCommandLine(String(info?.CommandLine ?? "")) ? "ours" : "other";
}

/**
 * What to do with a record whose process number is alive:
 * "discard" = the record is stale, forget it and start; "ready" = the service answered, nothing to do;
 * "kill" = our own process is stuck and the person asked for a repair; "fail" = our own process is stuck.
 * `waitHealthy` waits a while for the service to answer, because a slow start looks the same as a stuck one.
 */
export async function decideRecorded(owner, { repair = false, waitHealthy }) {
  if (owner === "other") return "discard";
  if (await waitHealthy()) return "ready";
  if (owner === "ours") return repair ? "kill" : "fail";
  return "discard"; // an unidentified program that never answered
}
