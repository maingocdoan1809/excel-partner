import type { ToolResult } from "../../../protocol/src/index.js";
import {
  TooLarge,
  fingerprintNow,
  fingerprintOf,
  locateInsert,
  photoCells,
  photograph,
  removeInsertedRows,
  restorePhoto,
  sameFingerprint,
  type Fingerprint,
  type InsertedRows,
  type Photo,
  type RestoreReport,
} from "./photo.js";
import { planUndo } from "./undo-plan.js";

/**
 * What the assistant changed, kept so it can be put back. Memory only: a reloaded task pane starts with an empty
 * journal, and the Changes list says so instead of offering an undo that would do nothing.
 */

export type EntryState = "undoable" | "undone" | "unavailable";

export interface JournalEntry {
  id: number;
  callId: string;
  requestId: string;
  tool: string;
  at: number;
  state: EntryState;
  /** Why it cannot be undone (state "unavailable"). */
  reason?: string;
  /** Things this change also did that undo does not put back. */
  caveats?: string[];
  photo?: Photo;
  after?: Fingerprint;
  inserted?: InsertedRows;
  cells: number;
  /** A copy of the whole workbook taken just before this reply's first change that could not be recorded. */
  backup?: BackupRef;
  backupError?: string;
}

export interface BackupRef {
  file: string;
  folder: string;
}

/** Saves a copy of the open workbook and says where it went. */
export type BackupWriter = () => Promise<BackupRef>;

export interface UndoOutcome {
  ok: boolean;
  message: string;
}

export const JOURNAL_LIMITS = { entries: 200, cells: 500_000 } as const;

/** The bridge labels tool calls made while no assistant reply is running "no-turn"; they belong to no reply. */
export const isReply = (requestId: string | undefined): requestId is string => !!requestId && requestId !== "no-turn";

type Execute = (tool: string, args: unknown) => Promise<ToolResult>;

const errorText = (error: unknown) => (error instanceof Error ? error.message : String(error));

export function describeRestore(report: RestoreReport): string {
  const parts = [`Restored ${report.restoredCells.toLocaleString()} cell${report.restoredCells === 1 ? "" : "s"}.`];
  if (report.missingCells > 0) parts.push(`${report.missingCells.toLocaleString()} cell${report.missingCells === 1 ? "" : "s"} could not be restored because their sheet was deleted.`);
  if (report.conflictCount > 0) {
    const where = report.conflicts.map((c) => `${c.sheet}!${c.address}`).join(", ");
    parts.push(`${report.conflictCount.toLocaleString()} cell${report.conflictCount === 1 ? " was" : "s were"} changed after the assistant's edit and left as they are (${where}${report.conflictCount > report.conflicts.length ? ", ..." : ""}).`);
  }
  if (report.unrestoredCount > 0) {
    const where = report.unrestored.map((c) => `${c.sheet}!${c.address}`).join(", ");
    parts.push(`${report.unrestoredCount.toLocaleString()} cell${report.unrestoredCount === 1 ? "" : "s"} could not be restored exactly (${where}${report.unrestoredCount > report.unrestored.length ? ", ..." : ""}).`);
  }
  parts.push(...report.notes.map((note) => `Note: ${note}.`));
  return parts.join(" ");
}

export class UndoJournal {
  private readonly entries: JournalEntry[] = [];
  private nextId = 1;
  private chain: Promise<unknown> = Promise.resolve();
  /** One workbook copy per assistant reply, taken lazily before the first change that cannot be undone. */
  private readonly backups = new Map<string, Promise<{ backup?: BackupRef; backupError?: string }>>();

  constructor(
    private readonly isWriteTool: (tool: string) => boolean,
    private readonly writeBackup?: BackupWriter,
  ) {}

  private backupFor(context: { callId: string; requestId: string }): Promise<{ backup?: BackupRef; backupError?: string }> {
    if (!this.writeBackup) return Promise.resolve({});
    // A call outside any reply must not share one copy with every later such call.
    const key = isReply(context.requestId) ? context.requestId : `call:${context.callId}`;
    let pending = this.backups.get(key);
    if (!pending) {
      pending = this.writeBackup().then(
        (backup) => ({ backup }),
        (error) => ({ backupError: errorText(error) }),
      );
      this.backups.set(key, pending);
      // Only the latest replies matter; the map must not grow for the life of the task pane.
      if (this.backups.size > 50) this.backups.delete(this.backups.keys().next().value!);
    }
    return pending;
  }

  /** One thing at a time touches the workbook: a photo is only true if nothing else changes the cells meanwhile. */
  private lock<T>(work: () => Promise<T>): Promise<T> {
    const run = this.chain.then(work, work);
    this.chain = run.catch(() => undefined);
    return run;
  }

  entriesFor(callId: string): JournalEntry[] {
    return this.entries.filter((entry) => entry.callId === callId);
  }

  entriesOf(requestId: string): JournalEntry[] {
    return this.entries.filter((entry) => entry.requestId === requestId);
  }

  /** Undo state of a recorded call: the most useful one among its entries. */
  stateFor(callId: string): { state: EntryState | "none"; reason?: string; caveats?: string[]; backup?: BackupRef; backupError?: string } {
    const mine = this.entriesFor(callId);
    if (mine.length === 0) return { state: "none" };
    const pick = mine.find((entry) => entry.state === "undoable") ?? mine.find((entry) => entry.state === "undone") ?? mine[0]!;
    return {
      state: pick.state,
      ...(pick.reason ? { reason: pick.reason } : {}),
      ...(pick.caveats ? { caveats: pick.caveats } : {}),
      ...(pick.backup ? { backup: pick.backup } : {}),
      ...(pick.backupError ? { backupError: pick.backupError } : {}),
    };
  }

  /** Run a tool, photographing what it is about to change when it is a workbook edit. */
  run(context: { callId: string; requestId: string }, tool: string, args: unknown, execute: Execute): Promise<ToolResult> {
    if (!this.isWriteTool(tool)) return execute(tool, args);
    return this.lock(() => this.record(context, tool, args, execute));
  }

  private add(context: { callId: string; requestId: string }, tool: string, fields: Partial<JournalEntry>): void {
    this.entries.push({ id: this.nextId++, callId: context.callId, requestId: context.requestId, tool, at: Date.now(), state: "undoable", cells: 0, ...fields });
    this.evict();
  }

  private evict(): void {
    let cells = this.entries.reduce((sum, entry) => sum + (entry.photo ? entry.cells : 0), 0);
    let kept = this.entries.filter((entry) => entry.photo || entry.inserted).length;
    for (const entry of this.entries) {
      if (cells <= JOURNAL_LIMITS.cells && kept <= JOURNAL_LIMITS.entries) break;
      if (!entry.photo && !entry.inserted) continue;
      cells -= entry.photo ? entry.cells : 0;
      kept--;
      delete entry.photo;
      delete entry.after;
      delete entry.inserted;
      if (entry.state === "undoable") {
        entry.state = "unavailable";
        entry.reason = "it is too old: the undo history only keeps the most recent changes";
      }
    }
  }

  private async record(context: { callId: string; requestId: string }, tool: string, args: unknown, execute: Execute): Promise<ToolResult> {
    const plan = planUndo(tool, args);
    if (plan.kind === "none") return execute(tool, args);
    if (plan.kind === "irreversible") {
      const saved = await this.backupFor(context);
      const result = await execute(tool, args);
      if (result.ok) this.add(context, tool, { state: "unavailable", reason: plan.reason, ...saved });
      return result;
    }
    if (plan.kind === "insert_rows") {
      let inserted: InsertedRows | null = null;
      try {
        inserted = await locateInsert(plan.sheet, plan.row, plan.count);
      } catch {
        inserted = null;
      }
      const result = await execute(tool, args);
      if (result.ok) {
        // The tool reports where it really inserted ("5:7"); trust that over the arguments.
        const reported = /^(\d+):(\d+)$/.exec(String((result.data as { shiftedRange?: unknown } | undefined)?.shiftedRange ?? ""));
        const where = reported ? { row: Number(reported[1]), count: Number(reported[2]) - Number(reported[1]) + 1 } : { row: plan.row, count: plan.count };
        this.add(context, tool, inserted ? { inserted: { ...inserted, ...where } } : { state: "unavailable", reason: "the sheet could not be identified" });
      }
      return result;
    }

    let photo: Photo | null = null;
    let reason = "";
    try {
      photo = await photograph(plan);
    } catch (error) {
      reason = error instanceof TooLarge ? `${error.message}, so it was not photographed` : `the cells could not be photographed (${errorText(error)})`;
    }
    const saved = photo ? {} : await this.backupFor(context);
    const result = await execute(tool, args);
    if (!photo) {
      if (result.ok) this.add(context, tool, { state: "unavailable", reason, ...saved });
      return result;
    }
    let after: Fingerprint;
    try {
      after = await fingerprintNow(photo);
    } catch (error) {
      if (result.ok) this.add(context, tool, { state: "unavailable", reason: `the result could not be recorded (${errorText(error)})` });
      return result;
    }
    // A refused call that changed nothing leaves no trace.
    if (!result.ok && sameFingerprint(after, fingerprintOf(photo))) return result;
    this.add(context, tool, { photo, after, cells: photoCells(photo), ...(plan.caveats ? { caveats: plan.caveats } : {}) });
    return result;
  }

  /** Undo every recorded edit of one tool call, newest first. */
  undoCall(callId: string): Promise<UndoOutcome> {
    return this.lock(() => this.undoEntries(this.entriesFor(callId)));
  }

  /** Undo every recorded edit of one assistant reply, newest first. */
  undoRequest(requestId: string): Promise<UndoOutcome> {
    return this.lock(() => this.undoEntries(this.entriesOf(requestId)));
  }

  private async undoEntries(all: JournalEntry[]): Promise<UndoOutcome> {
    const todo = all.filter((entry) => entry.state === "undoable").sort((a, b) => b.id - a.id);
    if (todo.length === 0) return { ok: false, message: "There is nothing to undo here." };
    const messages: string[] = [];
    let failed = false;
    for (const entry of todo) {
      try {
        if (entry.inserted) {
          const refusal = await removeInsertedRows(entry.inserted);
          if (refusal) {
            failed = true;
            messages.push(`Could not remove the inserted rows: ${refusal}.`);
            continue;
          }
          messages.push(`Removed ${entry.inserted.count} inserted row${entry.inserted.count === 1 ? "" : "s"}.`);
        } else if (entry.photo && entry.after) {
          const report = await restorePhoto(entry.photo, entry.after);
          messages.push(describeRestore(report));
          if (report.unrestoredCount > 0 || report.missingCells > 0) failed = true;
        } else {
          continue;
        }
        entry.state = "undone";
        delete entry.photo;
        delete entry.after;
        for (const caveat of entry.caveats ?? []) messages.push(`Note: ${caveat}.`);
      } catch (error) {
        failed = true;
        messages.push(`Could not undo ${entry.tool}: ${errorText(error)}`);
      }
    }
    return { ok: !failed, message: messages.join(" ") };
  }
}
