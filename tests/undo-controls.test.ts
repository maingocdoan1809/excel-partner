// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { lastUndoableReply, undoBar, undoControl, type UndoSource, type UndoView } from "../addin/taskpane/ui/undo-controls.js";
import type { StoredChange } from "../addin/taskpane/ui/sessions.js";

const change = (callId: string | undefined, requestId: string | undefined, ok = true): StoredChange => ({
  at: 1, tool: "write_range", ok, summary: "Wrote", target: null, ...(callId ? { callId } : {}), ...(requestId ? { requestId } : {}),
});

type Entry = { callId: string; requestId: string; state: "undoable" | "undone" | "unavailable"; reason?: string; caveats?: string[]; backup?: { file: string; folder: string }; backupError?: string };

function source(entries: Entry[]): UndoSource {
  return {
    entriesOf: (requestId) => entries.filter((e) => e.requestId === requestId),
    stateFor: (callId) => {
      const mine = entries.filter((e) => e.callId === callId);
      if (!mine.length) return { state: "none" };
      const { state, reason, caveats, backup, backupError } = mine[0]!;
      return { state, ...(reason ? { reason } : {}), ...(caveats ? { caveats } : {}), ...(backup ? { backup } : {}), ...(backupError ? { backupError } : {}) };
    },
  };
}

function view(entries: Entry[], extra: Partial<UndoView> = {}): UndoView & { undoCall: ReturnType<typeof vi.fn>; undoReply: ReturnType<typeof vi.fn> } {
  return { journal: source(entries), busy: false, note: null, undoCall: vi.fn(), undoReply: vi.fn(), ...extra } as never;
}

describe("undo controls in the Changes list", () => {
  it("offers Undo on a recorded edit and calls back with that edit's call", () => {
    const v = view([{ callId: "c1", requestId: "r1", state: "undoable" }]);
    const button = undoControl(change("c1", "r1"), v) as HTMLButtonElement;
    expect(button.tagName).toBe("BUTTON");
    expect(button.textContent).toBe("Undo");
    button.click();
    expect(v.undoCall).toHaveBeenCalledWith("c1");
  });

  it("lists what undo will not put back in the button's tooltip", () => {
    const v = view([{ callId: "c1", requestId: "r1", state: "undoable", caveats: ["cell notes are not restored"] }]);
    expect((undoControl(change("c1", "r1"), v) as HTMLElement).title).toBe("Note: cell notes are not restored.");
  });

  it("disables the buttons while an undo runs", () => {
    const v = view([{ callId: "c1", requestId: "r1", state: "undoable" }], { busy: true });
    expect((undoControl(change("c1", "r1"), v) as HTMLButtonElement).disabled).toBe(true);
    expect((undoBar([change("c1", "r1")], v)!.querySelector("button") as HTMLButtonElement).disabled).toBe(true);
  });

  it("says Undone after an undo, with no button", () => {
    const label = undoControl(change("c1", "r1"), view([{ callId: "c1", requestId: "r1", state: "undone" }]))!;
    expect(label.textContent).toBe("Undone");
    expect(label.tagName).toBe("SPAN");
  });

  it("explains an edit that cannot be undone, and names the workbook copy when there is one", () => {
    const plain = undoControl(change("c1", "r1"), view([{ callId: "c1", requestId: "r1", state: "unavailable", reason: "charts are not recorded" }]))!;
    expect(plain.textContent).toBe("Can't undo");
    expect(plain.title).toBe("Not recorded: charts are not recorded.");
    const withCopy = undoControl(change("c1", "r1"), view([{ callId: "c1", requestId: "r1", state: "unavailable", reason: "x", backup: { file: "Book1-1.xlsx", folder: "C:\\b" } }]))!;
    expect(withCopy.textContent).toBe("Can't undo · backup saved");
    expect(withCopy.title).toContain("Book1-1.xlsx in C:\\b");
    const failed = undoControl(change("c1", "r1"), view([{ callId: "c1", requestId: "r1", state: "unavailable", backupError: "bridge offline" }]))!;
    expect(failed.textContent).toBe("Can't undo");
    expect(failed.title).toContain("No copy of the workbook could be saved (bridge offline)");
  });

  it("tells the truth about edits from before a reload, and stays quiet about failed calls that changed nothing", () => {
    const empty = view([]);
    expect(undoControl(change("old", "r0"), empty)!.textContent).toBe("Can't undo");
    expect(undoControl(change("old", "r0"), empty)!.title).toContain("cleared when the task pane reloaded");
    expect(undoControl(change("old", "r0", false), empty)).toBeNull();
    expect(undoControl(change(undefined, undefined), empty)).toBeNull();
  });
});

describe("Undo last reply", () => {
  it("finds the newest reply that still has something to undo and counts its calls", () => {
    const entries: Entry[] = [
      { callId: "a", requestId: "r1", state: "undoable" },
      { callId: "b", requestId: "r2", state: "undoable" },
      { callId: "c", requestId: "r2", state: "undoable" },
      { callId: "d", requestId: "r3", state: "undone" },
    ];
    const changes = [change("a", "r1"), change("b", "r2"), change("c", "r2"), change("d", "r3")];
    expect(lastUndoableReply(changes, source(entries))).toEqual({ requestId: "r2", calls: 2 });
  });

  it("ignores calls made outside any reply", () => {
    const entries: Entry[] = [{ callId: "a", requestId: "no-turn", state: "undoable" }];
    expect(lastUndoableReply([change("a", "no-turn")], source(entries))).toBeNull();
    expect(undoBar([change("a", "no-turn")], view(entries))).toBeNull();
  });

  it("shows nothing when there is nothing to undo and no message", () => {
    expect(undoBar([change("a", "r1")], view([{ callId: "a", requestId: "r1", state: "undone" }]))).toBeNull();
  });

  it("offers one button for the whole reply and reports how many changes it covers", () => {
    const entries: Entry[] = [{ callId: "a", requestId: "r1", state: "undoable" }, { callId: "b", requestId: "r1", state: "undoable" }];
    const v = view(entries);
    const bar = undoBar([change("a", "r1"), change("b", "r1")], v)!;
    const button = bar.querySelector("button") as HTMLButtonElement;
    expect(button.textContent).toBe("Undo last reply (2 changes)");
    button.click();
    expect(v.undoReply).toHaveBeenCalledWith("r1");
    const single = undoBar([change("a", "r1")], view([entries[0]!]))!;
    expect(single.querySelector("button")!.textContent).toBe("Undo last reply (1 change)");
  });

  it("shows the outcome of the last undo, marking failures, even with nothing left to undo", () => {
    const failed = undoBar([change("a", "r1")], view([{ callId: "a", requestId: "r1", state: "undone" }], { note: { ok: false, text: "Could not undo write_range: lost focus" } }))!;
    const note = failed.querySelector(".undo-note") as HTMLElement;
    expect(note.textContent).toBe("Could not undo write_range: lost focus");
    expect(note.dataset.ok).toBe("false");
    expect(failed.querySelector("button")).toBeNull();
  });
});
