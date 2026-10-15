export type LogLevel = "debug" | "info" | "warn" | "error";

const ORDER: Record<LogLevel, number> = { debug: 10, info: 20, warn: 30, error: 40 };

export class Logger {
  constructor(
    private readonly scope: string,
    private readonly level: LogLevel = "info",
    private readonly sink: (line: string) => void = (line) => process.stderr.write(line + "\n"),
  ) {}

  child(scope: string): Logger {
    return new Logger(`${this.scope}:${scope}`, this.level, this.sink);
  }

  private emit(level: LogLevel, msg: string, extra?: unknown): void {
    if (ORDER[level] < ORDER[this.level]) return;
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
