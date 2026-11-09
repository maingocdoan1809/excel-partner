import type { AcceptRejectChangesInput, ReadTrackedChangesInput, SetTrackChangesInput } from "../../../../protocol/src/index.js";
import { WORD_COMMENTS, WORD_TRACK, isSupported } from "../../host/requirements.js";
import { ToolFailure } from "../common.js";
import { clip, fitResult, requireSupport, WORD_TRACK_MODE, wordRun } from "./common.js";

function dateText(value: unknown): string | undefined {
  if (value instanceof Date) return value.toISOString();
  return typeof value === "string" ? value : undefined;
}

const FALLBACK_TEXT_LIMIT = 8000;

/** Tracked changes (WordApi 1.6). Older Word only gets the original and current text, without who changed what. */
export async function readTrackedChanges(input: ReadTrackedChangesInput): Promise<unknown> {
  const max = input.max ?? 100;
  if (!isSupported(WORD_TRACK)) {
    // WordApi 1.4 can still give the two versions of the text.
    requireSupport("Reading individual tracked changes", WORD_COMMENTS, "Update Word, or read the text with read_doc_section.");
    return wordRun(async (context) => {
      const body = context.document.body;
      const original = body.getReviewedText("Original");
      const current = body.getReviewedText("Current");
      context.document.load("changeTrackingMode");
      await context.sync();
      const a = clip(original.value, FALLBACK_TEXT_LIMIT);
      const b = clip(current.value, FALLBACK_TEXT_LIMIT);
      return {
        summary: original.value === current.value ? "No tracked changes found (text only; this Word cannot list them)" : "The document has tracked changes (text only; this Word cannot list them)",
        limited: true,
        fallback: "reviewed_text",
        note: "This version of Word (below WordApi 1.6) cannot list individual tracked changes, authors or dates, nor accept or reject them through this add-in. Only the original and current text are available.",
        tracking_mode: context.document.changeTrackingMode,
        has_changes: original.value !== current.value,
        original_text: a.text,
        current_text: b.text,
        truncated: a.truncated || b.truncated,
      };
    });
  }
  return wordRun(async (context) => {
    const changes = context.document.body.getTrackedChanges();
    changes.load("items/author,items/date,items/text,items/type");
    context.document.load("changeTrackingMode");
    await context.sync();
    const items = changes.items;
    const rows = items.slice(0, max).map((c, i) => ({ index: i, type: c.type, author: c.author, date: dateText(c.date), text: clip(c.text, 300).text }));
    return fitResult(rows.length, (n) => ({
      tracking_mode: context.document.changeTrackingMode,
      total: items.length,
      returned: n,
      changes: rows.slice(0, n),
      truncated: n < items.length,
    }));
  });
}

export async function acceptRejectChanges(input: AcceptRejectChangesInput): Promise<unknown> {
  requireSupport("Accepting or rejecting tracked changes", WORD_TRACK, "Do it in Word's Review tab instead.");
  const accept = input.mode === "accept";
  const scope = input.scope ?? "all";
  return wordRun(async (context) => {
    const source = scope === "selection" ? context.document.getSelection().getTrackedChanges() : context.document.body.getTrackedChanges();
    source.load("items/author,items/type");
    await context.sync();
    const matching = input.author ? source.items.filter((c) => c.author === input.author) : source.items;
    if (matching.length === 0) {
      return {
        summary: `No tracked changes to ${input.mode}${input.author ? ` by ${input.author}` : ""}${scope === "selection" ? " in the selection" : ""}`,
        [accept ? "accepted" : "rejected"]: 0,
        scope,
      };
    }
    if (!input.author) {
      if (accept) source.acceptAll();
      else source.rejectAll();
    } else {
      for (const change of matching) {
        if (accept) change.accept();
        else change.reject();
      }
    }
    await context.sync();
    return {
      summary: `${accept ? "Accepted" : "Rejected"} ${matching.length} tracked change${matching.length === 1 ? "" : "s"}${input.author ? ` by ${input.author}` : ""}${scope === "selection" ? " in the selection" : ""}`,
      [accept ? "accepted" : "rejected"]: matching.length,
      scope,
    };
  });
}

export async function setTrackChanges(input: SetTrackChangesInput): Promise<unknown> {
  requireSupport("Track Changes", WORD_TRACK_MODE, "Turn it on in Word's Review tab.");
  return wordRun(async (context) => {
    context.document.load("changeTrackingMode");
    await context.sync();
    const previous = context.document.changeTrackingMode;
    const mode = input.enabled ? (input.mine_only ? "TrackMineOnly" : "TrackAll") : "Off";
    context.document.changeTrackingMode = mode;
    await context.sync();
    context.document.load("changeTrackingMode");
    await context.sync();
    if (context.document.changeTrackingMode !== mode) {
      throw new ToolFailure("office_error", `Word did not switch Track Changes to ${mode} (it is ${context.document.changeTrackingMode}). The document may be protected.`);
    }
    return {
      summary: input.enabled ? `Track Changes on${input.mine_only ? " (my edits only)" : ""}` : "Track Changes off",
      mode,
      previous_mode: previous,
    };
  });
}
