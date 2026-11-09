import type { ProposeDocEditsInput } from "../../../../protocol/src/index.js";
import { ToolFailure } from "../common.js";
import { clip, wordRun } from "./common.js";
import { replaceText } from "./edit.js";
import { locateText } from "./search.js";

/** One staged edit as the review card shows it. `index` is the position in the `edits` array the model sent. */
export interface ProposalItem {
  index: number;
  oldText: string;
  newText: string;
  reason?: string;
  alternatives?: string[];
  /** The old text was found (and has the requested occurrence). */
  found: boolean;
  /** The old text appears more than once and no occurrence was asked for, so the first match would be used. */
  ambiguous: boolean;
}

export interface ProposalRequest {
  title?: string;
  items: ProposalItem[];
}

export interface ProposalResponse {
  cancelled: boolean;
  /** Edits the person approved, with the wording they picked (newText or one of the alternatives). */
  accepted: Array<{ index: number; text: string }>;
}

export type ProposalReviewer = (request: ProposalRequest, signal?: AbortSignal) => Promise<ProposalResponse>;

let reviewer: ProposalReviewer | null = null;

/** The pane registers its review card here; with none registered, propose_doc_edits refuses to stage anything. */
export function setProposalReviewer(fn: ProposalReviewer | null): void {
  reviewer = fn;
}

export function hasProposalReviewer(): boolean {
  return reviewer !== null;
}

type Edit = ProposeDocEditsInput["edits"][number];

async function dryRun(edits: Edit[]): Promise<{ items: ProposalItem[]; problems: Map<number, string> }> {
  const problems = new Map<number, string>();
  const items = await wordRun(async (context) => {
    const out: ProposalItem[] = [];
    for (const [index, edit] of edits.entries()) {
      const occurrence = edit.occurrence ?? 1;
      let count = 0;
      try {
        count = (await locateText(context, context.document.body, edit.old_text, { matchCase: edit.match_case ?? true })).ranges.length;
      } catch (error) {
        if (!(error instanceof ToolFailure)) throw error;
        problems.set(index, error.message);
      }
      out.push({
        index,
        oldText: edit.old_text,
        newText: edit.new_text,
        ...(edit.reason !== undefined ? { reason: edit.reason } : {}),
        ...(edit.alternatives?.length ? { alternatives: edit.alternatives } : {}),
        found: count >= occurrence,
        ambiguous: count > 1 && occurrence === 1,
      });
    }
    return out;
  });
  return { items, problems };
}

const brief = (edit: Edit, text: string, index: number) => ({ index, old_text: clip(edit.old_text, 200).text, new_text: clip(text, 200).text });

/**
 * Stage edits for the person to review; nothing is written until they press Apply. Only approved edits are applied,
 * one by one in a single Word.run, each re-searching the document so an earlier edit cannot leave a stale position.
 */
export async function proposeDocEdits(input: ProposeDocEditsInput, signal?: AbortSignal): Promise<unknown> {
  const review = reviewer;
  if (!review) throw new ToolFailure("unsupported", "no review UI is available");
  const edits = input.edits;
  const { items, problems } = await dryRun(edits);
  const request: ProposalRequest = { ...(input.title !== undefined ? { title: input.title } : {}), items };
  const response = await review(request, signal);
  const notFound = items.filter((i) => !i.found).map((i) => ({ index: i.index, old_text: clip(i.oldText, 200).text, reason: problems.get(i.index) ?? "old_text was not found in the document" }));

  if (response.cancelled) {
    return {
      summary: "The person closed the review; nothing was changed",
      cancelled: true,
      applied: [],
      rejected: items.filter((i) => i.found).map((i) => ({ index: i.index, reason: "review cancelled" })),
      not_found: notFound,
    };
  }

  // Trust the card only for edits that exist and for wording the model offered.
  const chosen = new Map<number, string>();
  for (const pick of response.accepted) {
    const edit = edits[pick.index];
    if (!edit || chosen.has(pick.index)) continue;
    const allowed = [edit.new_text, ...(edit.alternatives ?? [])];
    if (allowed.includes(pick.text)) chosen.set(pick.index, pick.text);
  }
  const rejected = items.filter((i) => i.found && !chosen.has(i.index)).map((i) => ({ index: i.index, reason: "declined" }));

  const applied: Array<ReturnType<typeof brief>> = [];
  const lateNotFound: Array<{ index: number; old_text: string; reason: string }> = [];
  const state: { failure?: { index: number; message: string } } = {};
  if (chosen.size > 0) {
    await wordRun(async (context) => {
      for (const index of [...chosen.keys()].sort((a, b) => a - b)) {
        const edit = edits[index]!;
        const text = chosen.get(index)!;
        try {
          await replaceText(context, { oldText: edit.old_text, newText: text, matchCase: edit.match_case ?? true, occurrence: edit.occurrence ?? 1, all: false });
          applied.push(brief(edit, text, index));
        } catch (error) {
          if (error instanceof ToolFailure && (error.code === "range_not_found" || error.code === "invalid_arguments")) {
            lateNotFound.push({ index, old_text: clip(edit.old_text, 200).text, reason: error.message });
          } else {
            state.failure = { index, message: error instanceof Error ? error.message : String(error) };
            break;
          }
        }
      }
    });
  }
  const allNotFound = [...notFound, ...lateNotFound.filter((n) => !notFound.some((m) => m.index === n.index))];
  return {
    summary: `Applied ${applied.length} of ${edits.length} proposed edit${edits.length === 1 ? "" : "s"}${rejected.length ? `, ${rejected.length} declined` : ""}${allNotFound.length ? `, ${allNotFound.length} not found` : ""}`,
    cancelled: false,
    applied,
    rejected,
    not_found: allNotFound,
    ...(state.failure ? { stopped_at: state.failure, note: `Applying stopped at edit ${state.failure.index}: ${state.failure.message}. Later accepted edits were not applied.` } : {}),
  };
}
