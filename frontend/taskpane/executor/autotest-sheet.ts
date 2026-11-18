import type { AutotestEvent } from "../../../protocol/src/autotest.js";
import { recordTestResults } from "./test-results.js";

/** The open workbook as it is in Excel right now, unsaved edits included (a copy; the open file is not touched). */
export function exportWorkbook(): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (opened) => {
      if (opened.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(opened.error?.message ?? "Excel did not give the file."));
      const file = opened.value;
      const parts: number[][] = [];
      const fail = (message: string) => { file.closeAsync(); reject(new Error(message)); };
      const next = (index: number) => file.getSliceAsync(index, (slice) => {
        if (slice.status !== Office.AsyncResultStatus.Succeeded) return fail(slice.error?.message ?? "Excel stopped while giving the file.");
        parts.push(slice.value.data as number[]);
        if (index + 1 < file.sliceCount) return next(index + 1);
        file.closeAsync();
        const bytes = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
        let at = 0;
        for (const part of parts) { bytes.set(part, at); at += part.length; }
        resolve(bytes);
      });
      next(0);
    });
  });
}

const norm = (value: string) => value.replace(/\\/g, "/").toLowerCase();
const today = () => new Date().toISOString().slice(0, 10);

interface Cell { sheet: string; header: number; row: number; values: Record<string, string> }

/**
 * Turns the kit's progress events into the testcase's own cells, live. Cells for the same case are merged, and writes go out in
 * small batches (one Excel round trip per batch), so a fast run does not make Excel stutter.
 */
export class ProgressWriter {
  private queue = new Map<string, Cell>();
  private timer: ReturnType<typeof setTimeout> | null = null;
  private busy: Promise<void> = Promise.resolve();
  /** The copy of this workbook the kit reads; events about any other file are not ours to write. */
  file = "";

  constructor(
    private readonly resultValues: () => Record<string, string>,
    private readonly onError: (message: string) => void,
    private readonly delayMs = 700,
    private readonly write: typeof recordTestResults = recordTestResults,
  ) {}

  private text(status: string): string {
    return this.resultValues()[status] ?? (status === "RUNNING" ? "Đang chạy" : status);
  }

  /** True when the event changed (or will change) a cell of this workbook. */
  push(event: AutotestEvent): boolean {
    if (!this.file || !event.file || norm(event.file) !== norm(this.file) || !event.sheet || !event.row || !event.header) return false;
    const cols = event.cols ?? {};
    const result = cols.result ?? "Result";
    const values: Record<string, string> = {};
    if (event.kind === "case_started") values[result] = this.text("RUNNING");
    else if (event.kind === "step") values[result] = `${this.text("RUNNING")} ${event.n ?? ""}${event.total ? `/${event.total}` : ""}`.trim();
    else if (event.kind === "case_finished") {
      values[result] = this.text(event.status ?? "NOT EXECUTED");
      if (cols.date) values[cols.date] = today();
      if (cols.note && event.remark) values[cols.note] = event.remark.slice(0, 500);
    } else return false;
    const key = `${event.sheet}|${event.row}`;
    const cell = this.queue.get(key) ?? { sheet: event.sheet, header: event.header, row: event.row, values: {} };
    Object.assign(cell.values, values);
    this.queue.set(key, cell);
    // A finished case is written soon; progress ticks can wait for the batch.
    if (!this.timer) this.timer = setTimeout(() => { this.timer = null; void this.flush(); }, event.kind === "case_finished" ? 150 : this.delayMs);
    return true;
  }

  flush(): Promise<void> {
    this.busy = this.busy.then(async () => {
      const cells = [...this.queue.values()];
      this.queue.clear();
      const bySheet = new Map<string, Cell[]>();
      for (const cell of cells) bySheet.set(`${cell.sheet}|${cell.header}`, [...(bySheet.get(`${cell.sheet}|${cell.header}`) ?? []), cell]);
      for (const group of bySheet.values()) {
        try {
          await this.write({
            sheet: group[0]!.sheet, headerRow: group[0]!.header, addMissingColumns: true, allowOverwrite: true,
            results: group.slice(0, 200).map(cell => ({ row: cell.row, values: cell.values })),
          });
        } catch (error) { this.onError(error instanceof Error ? error.message : String(error)); }
      }
    });
    return this.busy;
  }
}
