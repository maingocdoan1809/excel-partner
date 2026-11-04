import { isReply, type BackupRef, type EntryState, type UndoOutcome } from "../executor/journal.js";
import type { StoredChange } from "./sessions.js";

/** What the Changes list needs from the undo journal. */
export interface UndoSource {
  entriesOf(requestId: string): Array<{ callId: string; state: EntryState }>;
  stateFor(callId: string): { state: EntryState | "none"; reason?: string; caveats?: string[]; backup?: BackupRef; backupError?: string };
}

export interface UndoView {
  journal: UndoSource;
  /** An undo is running: no second one may start. */
  busy: boolean;
  /** Result of the last undo, shown above the list until the next one. */
  note: { ok: boolean; text: string } | null;
  undoCall(callId: string): void;
  undoReply(requestId: string): void;
}

/** The newest reply that still has edits that can be put back, with how many tool calls that is. */
export function lastUndoableReply(changes: StoredChange[], journal: UndoSource): { requestId: string; calls: number } | null {
  for (let i = changes.length - 1; i >= 0; i--) {
    const requestId = changes[i]!.requestId;
    if (!isReply(requestId)) continue;
    const calls = new Set(journal.entriesOf(requestId).filter((entry) => entry.state === "undoable").map((entry) => entry.callId));
    if (calls.size > 0) return { requestId, calls: calls.size };
  }
  return null;
}

/** "Undo last reply" plus the outcome of the last undo; null when there is nothing to show. */
export function undoBar(changes: StoredChange[], view: UndoView): HTMLElement | null {
  const last = lastUndoableReply(changes, view.journal);
  if (!last && !view.note) return null;
  const bar = document.createElement("div");
  bar.className = "undo-bar";
  if (view.note) {
    const note = document.createElement("div");
    note.className = "undo-note";
    note.dataset.ok = String(view.note.ok);
    note.textContent = view.note.text;
    bar.append(note);
  }
  if (last) {
    const button = document.createElement("button");
    button.className = "secondary";
    button.disabled = view.busy;
    button.textContent = `Undo last reply (${last.calls} change${last.calls === 1 ? "" : "s"})`;
    button.title = "Put back everything the assistant changed in its latest reply. Cells you edited since then are left alone.";
    button.addEventListener("click", () => view.undoReply(last.requestId));
    bar.append(button);
  }
  return bar;
}

/** Undo button, or a quiet label saying why there is none. Failed calls that changed nothing have no entry and no label. */
export function undoControl(entry: StoredChange, view: UndoView): HTMLElement | null {
  if (!entry.callId) return null;
  const state = view.journal.stateFor(entry.callId);
  if (state.state === "none") {
    if (!entry.ok) return null;
    return undoLabel("Can't undo", "The undo history is kept in memory and was cleared when the task pane reloaded.");
  }
  if (state.state === "undone") return undoLabel("Undone");
  if (state.state === "unavailable") {
    const why = state.reason ? `Not recorded: ${state.reason}.` : "";
    const saved = state.backup
      ? ` A copy of the workbook from just before this reply was saved as ${state.backup.file} in ${state.backup.folder}; open it to go back.`
      : state.backupError ? ` No copy of the workbook could be saved (${state.backupError}).` : "";
    return undoLabel(state.backup ? "Can't undo · backup saved" : "Can't undo", `${why}${saved}`.trim() || undefined);
  }
  const button = document.createElement("button");
  button.className = "secondary";
  button.disabled = view.busy;
  button.textContent = "Undo";
  button.title = state.caveats?.length ? `Note: ${state.caveats.join("; ")}.` : "Put the cells back as they were before this change.";
  const callId = entry.callId;
  button.addEventListener("click", () => view.undoCall(callId));
  return button;
}

function undoLabel(text: string, title?: string): HTMLElement {
  const label = document.createElement("span");
  label.className = "undo-label";
  label.textContent = text;
  if (title) label.title = title;
  return label;
}

export type { UndoOutcome };
