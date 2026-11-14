import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { executeWordTool, setProposalReviewer } from "../frontend/taskpane/executor/word/index.js";
import type { ProposalRequest } from "../frontend/taskpane/executor/word/index.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";

let doc: FakeWordDoc;

beforeEach(() => {
  doc = installWordFake(new FakeWordDoc(["Revenue grew 5% in Q1.", "Costs were flat.", "Revenue is shown in dollars."]));
});
afterEach(() => {
  setProposalReviewer(null);
  uninstallWordFake();
});

async function data(tool: string, args: unknown = {}): Promise<any> {
  const result = await executeWordTool(tool, args);
  if (!result.ok) throw new Error(`${tool} failed: ${result.error.code}: ${result.error.message}`);
  return result.data;
}
async function failure(tool: string, args: unknown = {}) {
  const result = await executeWordTool(tool, args);
  expect(result.ok).toBe(false);
  return result.ok ? (undefined as never) : result.error;
}

describe("comments", () => {
  it("adds, reads, replies to and resolves a comment", async () => {
    const added = await data("add_comment", { anchor_text: "Costs were flat", text: "Is this right?" });
    expect(added).toMatchObject({ comment_id: expect.any(String), matches: 1 });
    const id = added.comment_id;

    const reply = await data("reply_comment", { comment_id: id, text: "Yes, checked." });
    expect(reply.reply_id).toEqual(expect.any(String));

    const read = await data("read_comments");
    expect(read).toMatchObject({ total: 1, unresolved: 1, returned: 1, truncated: false });
    expect(read.comments[0]).toMatchObject({ id, text: "Is this right?", anchor_text: "Costs were flat", resolved: false, author: "Tester" });
    expect(read.comments[0].replies).toEqual([expect.objectContaining({ text: "Yes, checked." })]);

    await data("resolve_comment", { comment_id: id });
    expect((await data("read_comments", { include_resolved: false })).comments).toEqual([]);
    expect((await data("read_comments")).comments[0].resolved).toBe(true);
    await data("resolve_comment", { comment_id: id, resolved: false });
    expect((await data("read_comments", { include_resolved: false })).comments).toHaveLength(1);
  });

  it("uses the requested occurrence for the anchor", async () => {
    await data("add_comment", { anchor_text: "Revenue", text: "second", occurrence: 2 });
    const read = await data("read_comments");
    expect(read.comments[0].anchor_text).toBe("Revenue");
    expect(doc.comments[0]!.anchor.p1).toBe(doc.paragraphs[2]);
  });

  it("limits the number returned", async () => {
    for (let i = 0; i < 3; i++) await data("add_comment", { anchor_text: "Costs", text: `c${i}` });
    const read = await data("read_comments", { max: 2 });
    expect(read).toMatchObject({ total: 3, returned: 2, truncated: true });
  });

  it("fails clearly for an unknown comment id", async () => {
    const error = await failure("reply_comment", { comment_id: "nope", text: "x" });
    expect(error.code).toBe("range_not_found");
    expect(error.message).toContain("read_comments");
  });

  it("loses the comment when its anchored text is replaced, which is why callers reply first", async () => {
    const { comment_id } = await data("add_comment", { anchor_text: "Costs were flat", text: "?" });
    await data("reply_comment", { comment_id, text: "answered" });
    await data("resolve_comment", { comment_id });
    expect(doc.comments[0]!.resolved).toBe(true);
    await data("edit_doc_text", { old_text: "Costs were flat", new_text: "Costs fell" });
    expect(doc.comments).toHaveLength(0);
  });

  it("is unsupported below WordApi 1.4 for every comment tool", async () => {
    installWordFake(doc, { sets: { WordApi: "1.3" } });
    for (const [tool, args] of [["read_comments", {}], ["add_comment", { anchor_text: "Costs", text: "x" }], ["reply_comment", { comment_id: "1", text: "x" }], ["resolve_comment", { comment_id: "1" }]] as const) {
      const error = await failure(tool, args);
      expect(error.code, tool).toBe("unsupported");
      expect(error.message).toContain("WordApi 1.4");
    }
  });
});

describe("tracked changes", () => {
  it("reads tracked changes after a tracked edit", async () => {
    await data("set_track_changes", { enabled: true });
    await data("edit_doc_text", { old_text: "flat", new_text: "down" });
    const read = await data("read_tracked_changes");
    expect(read).toMatchObject({ tracking_mode: "TrackAll", total: 2 });
    expect(read.changes.map((c: any) => [c.type, c.text, c.author])).toEqual([["Deleted", "flat", "Tester"], ["Added", "down", "Tester"]]);
  });

  it("accepts all, rejects by author, and reports an empty result", async () => {
    doc.tracked.push({ type: "Added", author: "Ann", date: new Date(), text: "a" }, { type: "Deleted", author: "Bob", date: new Date(), text: "b" });
    const byAuthor = await data("accept_reject_changes", { mode: "reject", author: "Bob" });
    expect(byAuthor).toMatchObject({ rejected: 1, scope: "all" });
    expect(doc.tracked.map((c) => c.author)).toEqual(["Ann"]);
    const all = await data("accept_reject_changes", { mode: "accept" });
    expect(all.accepted).toBe(1);
    expect(doc.tracked).toHaveLength(0);
    expect(await data("accept_reject_changes", { mode: "accept" })).toMatchObject({ accepted: 0 });
  });

  it("limits to the selection with scope selection", async () => {
    doc.tracked.push({ type: "Added", author: "Ann", date: new Date(), text: "a" });
    expect(await data("accept_reject_changes", { mode: "accept", scope: "selection" })).toMatchObject({ accepted: 1, scope: "selection" });
  });

  it("sets and clears the tracking mode", async () => {
    expect(await data("set_track_changes", { enabled: true, mine_only: true })).toMatchObject({ mode: "TrackMineOnly", previous_mode: "Off" });
    expect(doc.changeTrackingMode).toBe("TrackMineOnly");
    expect(await data("set_track_changes", { enabled: false })).toMatchObject({ mode: "Off" });
  });

  it("falls back to reviewed text on WordApi 1.4/1.5 and flags it", async () => {
    await data("set_track_changes", { enabled: true });
    await data("edit_doc_text", { old_text: "flat", new_text: "down" });
    installWordFake(doc, { sets: { WordApi: "1.5" } });
    const read = await data("read_tracked_changes");
    expect(read).toMatchObject({ limited: true, fallback: "reviewed_text", has_changes: true });
    expect(read.original_text).toContain("flat");
    expect(read.current_text).toContain("down");
  });

  it("is unsupported below 1.4 for reading and below 1.6 for accept/reject", async () => {
    installWordFake(doc, { sets: { WordApi: "1.3" } });
    expect((await failure("read_tracked_changes")).code).toBe("unsupported");
    expect((await failure("set_track_changes", { enabled: true })).code).toBe("unsupported");
    installWordFake(doc, { sets: { WordApi: "1.5" } });
    expect((await failure("accept_reject_changes", { mode: "accept" })).message).toContain("WordApi 1.6");
  });
});

describe("propose_doc_edits", () => {
  const edits = [
    { old_text: "grew 5%", new_text: "rose 5%", reason: "tone", alternatives: ["increased 5%"] },
    { old_text: "Costs were flat.", new_text: "Costs held steady." },
    { old_text: "Revenue", new_text: "Sales" }, // appears twice: ambiguous
    { old_text: "does not exist", new_text: "x" },
  ];

  it("refuses when no review UI is registered and writes nothing", async () => {
    const error = await failure("propose_doc_edits", { edits });
    expect(error).toMatchObject({ code: "unsupported", message: "no review UI is available" });
    expect(doc.texts()[0]).toBe("Revenue grew 5% in Q1.");
  });

  it("shows found / ambiguous per edit, writes nothing before the reviewer answers, then applies only accepted edits", async () => {
    let seen: ProposalRequest | undefined;
    let textDuringReview: string[] = [];
    setProposalReviewer(async (request) => {
      seen = request;
      textDuringReview = doc.texts();
      return { cancelled: false, accepted: [{ index: 0, text: "increased 5%" }, { index: 2, text: "Sales" }] };
    });
    const out = await data("propose_doc_edits", { title: "Tone", edits });
    expect(textDuringReview).toEqual(["Revenue grew 5% in Q1.", "Costs were flat.", "Revenue is shown in dollars."]);
    expect(seen!.title).toBe("Tone");
    expect(seen!.items.map((i) => [i.index, i.found, i.ambiguous])).toEqual([[0, true, false], [1, true, false], [2, true, true], [3, false, false]]);
    expect(seen!.items[0]).toMatchObject({ oldText: "grew 5%", newText: "rose 5%", reason: "tone", alternatives: ["increased 5%"] });
    // edit 2 ("Revenue", occurrence 1) hits the first "Revenue", which is why the card flags it as ambiguous
    expect(doc.texts()).toEqual(["Sales increased 5% in Q1.", "Costs were flat.", "Revenue is shown in dollars."]);
    expect(out.applied.map((a: any) => [a.index, a.new_text])).toEqual([[0, "increased 5%"], [2, "Sales"]]);
    expect(out.rejected).toEqual([{ index: 1, reason: "declined" }]);
    expect(out.not_found).toEqual([expect.objectContaining({ index: 3 })]);
    expect(out.cancelled).toBe(false);
  });

  it("applies several edits in one Word.run, re-searching so earlier edits cannot break later ones", async () => {
    let runs = 0;
    const original = (globalThis as any).Word.run;
    (globalThis as any).Word.run = (cb: any) => { runs++; return original(cb); };
    setProposalReviewer(async () => ({ cancelled: false, accepted: [{ index: 0, text: "rose 5%" }, { index: 1, text: "Costs held steady." }] }));
    await data("propose_doc_edits", { edits: edits.slice(0, 2) });
    expect(runs).toBe(2); // one dry run, one apply
    expect(doc.texts().slice(0, 2)).toEqual(["Revenue rose 5% in Q1.", "Costs held steady."]);
  });

  it("changes nothing when the person closes the review", async () => {
    setProposalReviewer(async () => ({ cancelled: true, accepted: [] }));
    const out = await data("propose_doc_edits", { edits });
    expect(out).toMatchObject({ cancelled: true, applied: [] });
    expect(out.rejected).toHaveLength(3);
    expect(doc.texts()[0]).toBe("Revenue grew 5% in Q1.");
  });

  it("applies nothing when the person rejects everything", async () => {
    setProposalReviewer(async () => ({ cancelled: false, accepted: [] }));
    const out = await data("propose_doc_edits", { edits });
    expect(out).toMatchObject({ cancelled: false, applied: [] });
    expect(out.rejected).toHaveLength(3);
  });

  it("ignores accepted wording the model never offered and indexes that do not exist", async () => {
    setProposalReviewer(async () => ({ cancelled: false, accepted: [{ index: 0, text: "something sneaky" }, { index: 99, text: "x" }] }));
    const out = await data("propose_doc_edits", { edits });
    expect(out.applied).toEqual([]);
    expect(doc.texts()[0]).toBe("Revenue grew 5% in Q1.");
  });

  it("reports an accepted edit whose text vanished meanwhile as not found", async () => {
    setProposalReviewer(async () => {
      doc.paragraphs[1]!.text = "Costs changed under the review.";
      return { cancelled: false, accepted: [{ index: 1, text: "Costs held steady." }] };
    });
    const out = await data("propose_doc_edits", { edits });
    expect(out.applied).toEqual([]);
    expect(out.not_found.map((n: any) => n.index)).toContain(1);
  });

  it("passes the abort signal to the reviewer", async () => {
    const controller = new AbortController();
    let received: AbortSignal | undefined;
    setProposalReviewer(async (_request, signal) => {
      received = signal;
      return { cancelled: true, accepted: [] };
    });
    await executeWordTool("propose_doc_edits", { edits }, controller.signal);
    expect(received).toBe(controller.signal);
  });
});
