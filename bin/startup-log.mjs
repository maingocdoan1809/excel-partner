// What happened the last times Office Partner was started at sign-in or by "Fix connection".
// Windows runs the sign-in entry with no window, so without this file a failed start leaves no trace at all.
// Importing this file needs only Node built-ins (it is also copied into the recovery launcher).
// The doctor reads this file with its own small reader (readLastStartup in doctor-probes.mjs) so that it stays standalone.
import fs from "node:fs";
import path from "node:path";

const FILE = "startup.log";
const MAX_BYTES = 32 * 1024;
const KEEP_LINES = 40;

/** A stable code for the failure text, so diagnostics can name the problem without copying paths into a report. */
export function startupCode(message) {
  const text = String(message ?? "");
  if (/setup'? first|INSTALLATION_NOT_INITIALIZED/i.test(text)) return "NOT_SET_UP";
  if (/RELEASE_DAMAGED|RECOVERY_|ROLLBACK_/i.test(text)) return "RELEASE_DAMAGED";
  if (/not answering|still alive but unhealthy|not healthy/i.test(text)) return "NOT_ANSWERING";
  if (/in use|EADDRINUSE|occupied/i.test(text)) return "PORT_IN_USE";
  if (/did not start|failed to start|failed to listen/i.test(text)) return "START_FAILED";
  if (/certificate|TLS/i.test(text)) return "CERTIFICATE_PROBLEM";
  return "START_FAILED";
}

/** Append one start attempt: { service, attempt, ok, code?, message? }. Never throws: logging must not stop a start. */
export function recordStartup(dataDir, entry) {
  try {
    const file = path.join(dataDir, FILE);
    fs.mkdirSync(dataDir, { recursive: true });
    const line = JSON.stringify({ at: new Date().toISOString(), ...entry, message: entry.message ? String(entry.message).split(/\r?\n/)[0].slice(0, 300) : undefined });
    try {
      if (fs.statSync(file).size > MAX_BYTES) {
        const kept = fs.readFileSync(file, "utf8").split("\n").filter(Boolean).slice(-KEEP_LINES);
        fs.writeFileSync(file, kept.join("\n") + "\n", { mode: 0o600 });
      }
    } catch { /* No file yet. */ }
    fs.appendFileSync(file, line + "\n", { mode: 0o600 });
  } catch { /* A read-only or missing data folder must not make the start fail. */ }
}
