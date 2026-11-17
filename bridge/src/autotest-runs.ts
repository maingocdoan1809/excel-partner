import fs from "node:fs";
import path from "node:path";
import { type AutotestEvent, type AutotestRunState, type RunPhase } from "../../protocol/src/autotest.js";
import { parseEvents, reduceRun } from "../../protocol/src/autotest-state.js";

export { parseEvents, reduceRun, STALE_RUN_MS } from "../../protocol/src/autotest-state.js";

export interface RunListing { id: string; phase: RunPhase; started: string; updated: string; counts: Record<string, number>; total: number; stopReason: string }
export type RunListener = (run: string, event: AutotestEvent) => void;

/** The runs under `<autotest home>/test_execution_output`: `runs/<run id>/events.jsonl` and `adhoc/<date>/events.jsonl`. */
export class AutotestRuns {
  constructor(private readonly where: string | (() => string)) {}

  private get outDir(): string { return typeof this.where === "string" ? this.where : this.where(); }

  private files(): Array<{ id: string; file: string }> {
    const found: Array<{ id: string; file: string }> = [];
    for (const group of ["runs", "adhoc"]) {
      let names: string[] = [];
      try { names = fs.readdirSync(path.join(this.outDir, group)); } catch { continue; }
      for (const name of names) {
        const file = path.join(this.outDir, group, name, "events.jsonl");
        if (fs.existsSync(file)) found.push({ id: group === "runs" ? name : `adhoc-${name}`, file });
      }
    }
    return found;
  }

  private fileOf(id: string): string | null {
    return this.files().find(entry => entry.id === id)?.file ?? null;
  }

  events(id: string, from = 0): { events: AutotestEvent[]; next: number } | null {
    const file = this.fileOf(id);
    if (!file) return null;
    const all = parseEvents(fs.readFileSync(file, "utf8"));
    return { events: all.slice(from), next: all.length };
  }

  state(id: string): AutotestRunState | null {
    const read = this.events(id);
    return read ? reduceRun(id, read.events) : null;
  }

  list(limit = 30): RunListing[] {
    return this.files()
      .map(entry => ({ ...entry, mtime: fs.statSync(entry.file).mtimeMs }))
      .sort((a, b) => b.mtime - a.mtime)
      .slice(0, limit)                                             // only the newest runs are read, however many exist
      .map(({ id, file }) => {
        const state = reduceRun(id, parseEvents(fs.readFileSync(file, "utf8")));
        return { id, phase: state.phase, started: state.started, updated: state.updated, counts: state.counts, total: state.cases.length, stopReason: state.stopReason };
      });
  }

  /**
   * The file behind an image path written in an event, only when it really is inside this output folder. A relative path (a run started with a
   * relative QA_OUT) is relative to the autotest folder, the parent of the output folder, not to wherever the bridge was started.
   */
  image(id: string, requested: string): string | null {
    if (!this.fileOf(id)) return null;
    try {
      const root = fs.realpathSync(this.outDir);
      const real = fs.realpathSync(path.isAbsolute(requested) ? requested : path.resolve(path.dirname(this.outDir), requested));
      // Windows paths differ by case and still name the same file
      const same = (value: string) => process.platform === "win32" ? value.toLowerCase() : value;
      return same(real).startsWith(same(root) + path.sep) && /\.(png|jpe?g)$/i.test(real) ? real : null;
    } catch { return null; }
  }

  /** Call `listener` for every event written from now on. Files that already exist are not replayed. */
  watch(listener: RunListener, intervalMs = 1000): () => void {
    const offsets = new Map<string, number>();
    for (const { file } of this.files()) offsets.set(file, fs.statSync(file).size);
    const tick = () => {
      for (const { id, file } of this.files()) {
        let size = 0;
        try { size = fs.statSync(file).size; } catch { continue; }
        const from = offsets.get(file) ?? 0;
        if (size < from) { offsets.set(file, 0); continue; }
        if (size === from) continue;
        const fd = fs.openSync(file, "r");
        try {
          const buffer = Buffer.alloc(size - from);
          fs.readSync(fd, buffer, 0, buffer.length, from);
          const text = buffer.toString("utf8");
          const complete = text.lastIndexOf("\n") + 1;
          offsets.set(file, from + Buffer.byteLength(text.slice(0, complete)));
          for (const event of parseEvents(text.slice(0, complete))) listener(id, event);
        } finally { fs.closeSync(fd); }
      }
    };
    const timer = setInterval(() => { try { tick(); } catch { /* retried on the next tick */ } }, intervalMs);
    timer.unref();
    return () => clearInterval(timer);
  }
}
