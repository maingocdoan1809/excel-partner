import { createHash } from "node:crypto";

/** How long an identical call stays blocked after Excel dropped while running it. */
export const CRASH_BLOCK_MS = 15 * 60_000;

function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map(key => `${JSON.stringify(key)}:${stable((value as Record<string, unknown>)[key])}`).join(",")}}`;
  }
  return JSON.stringify(value) ?? "null";
}

/**
 * Remembers tool calls that were still running when the task pane dropped, which almost always means Excel crashed
 * on them. Excel kills the whole process, so the add-in cannot catch it; the only defence is not sending the same
 * call again. The memory is shared by every session because Excel's restart opens a new one.
 */
export class CrashGuard {
  private readonly blocked = new Map<string, number>();

  constructor(private readonly ttlMs = CRASH_BLOCK_MS, private readonly now: () => number = Date.now) {}

  private key(tool: string, args: unknown): string {
    return `${tool}:${createHash("sha1").update(stable(args)).digest("hex")}`;
  }

  record(tool: string, args: unknown): void {
    this.blocked.set(this.key(tool, args), this.now() + this.ttlMs);
  }

  /** Minutes left for which this exact call stays blocked, or 0 when it may run. */
  blockedMinutes(tool: string, args: unknown): number {
    const key = this.key(tool, args);
    const until = this.blocked.get(key);
    if (until === undefined) return 0;
    const left = until - this.now();
    if (left <= 0) { this.blocked.delete(key); return 0; }
    return Math.ceil(left / 60_000);
  }
}

export function crashBlockedMessage(tool: string, minutes: number): string {
  return `Blocked: this exact ${tool} call was running when Excel dropped the connection, so Excel most likely crashed on it. `
    + `Repeating it would crash Excel again, so it stays blocked for about ${minutes} more minute(s). `
    + "Do not retry it. Tell the user, then use a different approach (another tool, one sheet, or a smaller range).";
}
