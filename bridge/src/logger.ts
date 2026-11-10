import fs from "node:fs";
import path from "node:path";
import os from "node:os";

function writeLog(line: string) {
  const file = process.env.OFFICE_PARTNER_LOG_FILE;
  if (!file) { process.stderr.write(line + "\n"); return; }
  try {
    if (fs.existsSync(file) && fs.statSync(file).size > 5 * 1024 * 1024) {
      for (let index = 3; index >= 1; index--) {
        const source = index === 1 ? file : `${file}.${index - 1}`, target = `${file}.${index}`;
        if (fs.existsSync(target)) fs.unlinkSync(target);
        if (fs.existsSync(source)) fs.renameSync(source, target);
      }
    }
    fs.appendFileSync(file, line + "\n", { mode: 0o600 });
  } catch { process.stderr.write(line + "\n"); }
}

export type LogLevel = "debug" | "info" | "warn" | "error";

const ORDER: Record<LogLevel, number> = { debug: 10, info: 20, warn: 30, error: 40 };

export class Logger {
  constructor(
    private readonly scope: string,
    private readonly level: LogLevel = "info",
    private readonly sink: (line: string) => void = writeLog,
  ) {}

  child(scope: string): Logger {
    return new Logger(`${this.scope}:${scope}`, this.level, this.sink);
  }

  private emit(level: LogLevel, msg: string, extra?: unknown): void {
    if (ORDER[level] < ORDER[debugging() ? "debug" : this.level]) return;
    const ts = new Date().toISOString();
    const base = `${ts} ${level.toUpperCase().padEnd(5)} [${this.scope}] ${msg}`;
    this.sink(extra === undefined ? base : `${base} ${safeJson(extra)}`);
  }

  debug(msg: string, extra?: unknown): void {
    this.emit("debug", msg, extra);
  }
  info(msg: string, extra?: unknown): void {
    this.emit("info", msg, extra);
  }
  warn(msg: string, extra?: unknown): void {
    this.emit("warn", msg, extra);
  }
  error(msg: string, extra?: unknown): void {
    this.emit("error", msg, extra);
  }
}

function safeJson(value: unknown): string {
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}
let debugReadAt = 0, debugUntil = 0;
function debugging() {
  if (Date.now() - debugReadAt > 1000) {
    debugReadAt = Date.now();
    try {
      const file = path.join(process.env.OFFICE_PARTNER_DATA_DIR || path.join(process.env.LOCALAPPDATA || os.homedir(), "Excel Partner"), "debug.json");
      const value = JSON.parse(fs.readFileSync(file, "utf8")).expires;
      debugUntil = typeof value === "number" ? Math.min(value, Date.now() + 15 * 60000) : 0;
    } catch { debugUntil = 0; }
  }
  return Date.now() < debugUntil;
}
