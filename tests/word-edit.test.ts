import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { executeWordTool } from "../frontend/taskpane/executor/word/index.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";

let doc: FakeWordDoc;

beforeEach(() => {
  doc = installWordFake(new FakeWordDoc(["The quick brown fox.", "The lazy dog sleeps.", "The end."]));
});
afterEach(() => uninstallWordFake());

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

describe("edit_doc_text", () => {
  it("replaces the first match by default and says how many there were", async () => {
    const out = await data("edit_doc_text", { old_text: "The", new_text: "A", match_case: true });
    expect(doc.texts()).toEqual(["A quick brown fox.", "The lazy dog sleeps.", "The end."]);
    expect(out).toMatchObject({ replaced: 1, matches: 3, occurrence: 1 });
    expect(out.summary).toContain("Replaced 1 match");
    expect(out.note).toContain("occurrence");
  });

  it("replaces the requested occurrence", async () => {
    await data("edit_doc_text", { old_text: "The", new_text: "That", occurrence: 3 });
    expect(doc.texts()).toEqual(["The quick brown fox.", "The lazy dog sleeps.", "That end."]);
  });

  it("reports not found, with the case-sensitivity hint", async () => {
    const error = await failure("edit_doc_text", { old_text: "the", new_text: "x" });
    expect(error.code).toBe("range_not_found");
    expect(error.message).toContain("case-sensitive");
    expect(doc.texts()[0]).toBe("The quick brown fox.");
    const loose = await data("edit_doc_text", { old_text: "THE", new_text: "x", match_case: false });
    expect(loose.replaced).toBe(1);
  });

  it("fails when the occurrence does not exist", async () => {
    const error = await failure("edit_doc_text", { old_text: "The", new_text: "x", occurrence: 9 });
    expect(error.code).toBe("range_not_found");
    expect(error.message).toContain("3 times");
  });

  it("replaces every match from one search with all", async () => {
    let searches = 0;
    const original = (globalThis as any).Word.run;
    (globalThis as any).Word.run = (cb: any) => original((ctx: any) => {
      const body = ctx.document.body;
      const search = body.search;
      body.search = (...a: unknown[]) => { searches++; return search(...a); };
      return cb(ctx);
    });
    const out = await data("edit_doc_text", { old_text: "The", new_text: "A", all: true });
    expect(doc.texts()).toEqual(["A quick brown fox.", "A lazy dog sleeps.", "A end."]);
    expect(out).toMatchObject({ replaced: 3, matches: 3 });
    expect(searches).toBe(1);
  });

  it("deletes the match when new_text is empty and keeps the other paragraphs", async () => {
    await data("edit_doc_text", { old_text: " brown", new_text: "" });
    expect(doc.texts()).toEqual(["The quick fox.", "The lazy dog sleeps.", "The end."]);
  });

  it("matches text containing ^ literally and turns \\n into a paragraph break", async () => {
    doc.setParagraphs(["x ^ y", "next"]);
    await data("edit_doc_text", { old_text: "x ^ y", new_text: "x^2" });
    expect(doc.texts()).toEqual(["x^2", "next"]);
    await data("edit_doc_text", { old_text: "x^2\nnext", new_text: "joined" });
    expect(doc.texts()).toEqual(["joined"]);
    await data("edit_doc_text", { old_text: "joined", new_text: "a\nb" });
    expect(doc.texts()).toEqual(["a", "b"]);
  });

  it("finds old_text longer than 255 characters through its start and end", async () => {
    const long = `${"alpha ".repeat(30)}MIDDLE${" omega".repeat(30)}`;
    expect(long.length).toBeGreaterThan(255);
    doc.setParagraphs(["before", long, "after"]);
    const out = await data("edit_doc_text", { old_text: long, new_text: "short", match_case: true });
    expect(doc.texts()).toEqual(["before", "short", "after"]);
    expect(out.matched_by).toContain("start_and_end");
  });

  it("refuses a long old_text whose ends exist but whose middle differs", async () => {
    const long = `${"alpha ".repeat(30)}MIDDLE${" omega".repeat(30)}`;
    doc.setParagraphs([long.replace("MIDDLE", "OTHER")]);
    const error = await failure("edit_doc_text", { old_text: long, new_text: "x" });
    expect(error.code).toBe("range_not_found");
    expect(doc.texts()[0]).toContain("OTHER");
  });

  it("picks the right one of several long matches by their text", async () => {
    const head = "alpha ".repeat(30);
    const tail = " omega".repeat(30);
    doc.setParagraphs([`${head}ONE${tail}`, `${head}TWO${tail}`]);
    await data("edit_doc_text", { old_text: `${head}TWO${tail}`, new_text: "replaced" });
    expect(doc.texts()).toEqual([`${head}ONE${tail}`, "replaced"]);
  });

  it("turns Track Changes on only when asked and supported", async () => {
    const out = await data("edit_doc_text", { old_text: "quick", new_text: "slow", track: true });
    expect(doc.changeTrackingMode).toBe("TrackAll");
    expect(out).toMatchObject({ tracked: true, track_changes_left_on: true });
    expect(doc.tracked.map((c) => c.type)).toEqual(["Deleted", "Added"]);
  });

  it("does not touch Track Changes when track is omitted", async () => {
    await data("edit_doc_text", { old_text: "quick", new_text: "slow" });
    expect(doc.changeTrackingMode).toBe("Off");
  });

  it("refuses track:true on Word below 1.6 instead of editing untracked", async () => {
    installWordFake(doc, { sets: { WordApi: "1.4" } });
    const error = await failure("edit_doc_text", { old_text: "quick", new_text: "slow", track: true });
    expect(error.code).toBe("unsupported");
    expect(error.message).toContain("WordApi 1.6");
    expect(doc.texts()[0]).toBe("The quick brown fox.");
  });

  it("rejects bad arguments before touching Word", async () => {
    expect((await failure("edit_doc_text", { old_text: "", new_text: "x" })).code).toBe("invalid_arguments");
    expect((await failure("edit_doc_text", {})).code).toBe("invalid_arguments");
  });

  it("returns unknown_tool for a name that is not a Word tool", async () => {
    expect((await failure("write_range", {})).code).toBe("unknown_tool");
  });

  it("stops with unsupported when Word is not there at all", async () => {
    uninstallWordFake();
    expect((await failure("get_doc_outline")).code).toBe("unsupported");
  });
});

describe("edit_doc_list", () => {
  it("inserts a bulleted list with the List API after a paragraph", async () => {
    const out = await data("edit_doc_list", { operation: "bullets", items: ["one", { text: "two", level: 1 }, "three"], after_paragraph: 0 });
    expect(doc.texts()).toEqual(["The quick brown fox.", "one", "two", "three", "The lazy dog sleeps.", "The end."]);
    expect(out).toMatchObject({ items: 3, first_paragraph_index: 1, method: "list_api", after_paragraph: 0 });
    const [, a, b, c] = doc.paragraphs;
    expect([a!.listId, b!.listId, c!.listId].every((id) => id !== null && id === a!.listId)).toBe(true);
    expect([a!.level, b!.level, c!.level]).toEqual([0, 1, 0]);
    expect(a!.styleBuiltIn).toBe("ListParagraph");
    expect(doc.lists.get(a!.listId!)!.bullets).toEqual({ 0: "Solid", 1: "Hollow" });
  });

  it("numbers a list at the end of the document", async () => {
    const out = await data("edit_doc_list", { operation: "numbered", items: ["first", "second"] });
    expect(doc.texts().slice(-2)).toEqual(["first", "second"]);
    expect(out.first_paragraph_index).toBe(3);
    expect(doc.lists.get(doc.paragraphs.at(-1)!.listId!)!.numbering).toEqual({ 0: "Arabic" });
  });

  it("falls back to HTML on Word below 1.3", async () => {
    installWordFake(doc, { sets: { WordApi: "1.2" } });
    const out = await data("edit_doc_list", { operation: "numbered", items: ["a", { text: "b", level: 1 }, "c <x> & y"] });
    expect(out.method).toBe("html");
    expect(doc.htmlInserts[0]!.html).toBe("<ol><li>a<ol><li>b</li></ol></li><li>c &lt;x&gt; &amp; y</li></ol>");
  });

  it("rejects an after_paragraph outside the document", async () => {
    expect((await failure("edit_doc_list", { operation: "bullets", items: ["x"], after_paragraph: 50 })).code).toBe("invalid_arguments");
  });
});

describe("collapse_blank_paragraphs", () => {
  beforeEach(() => {
    doc.setParagraphs([
      "start", "", " ", "\u00a0", "text",
      { text: "", pictures: 1 }, "", "",
      "\f", "", "", "",
      { text: "", tableNestingLevel: 1 }, { text: "", tableNestingLevel: 1 }, { text: "", tableNestingLevel: 1 },
      "tail", "", "",
    ]);
  });

  it("previews without changing anything", async () => {
    const out = await data("collapse_blank_paragraphs", { max: 1, preview: true });
    expect(doc.paragraphs).toHaveLength(18);
    expect(out).toMatchObject({ preview: true, would_remove: 2 + 1 + 2 + 1 });
    expect(out.runs.map((r: any) => r.start_index)).toEqual([1, 6, 9, 16]);
  });

  it("keeps page breaks, pictures, table cells and the last paragraph", async () => {
    const out = await data("collapse_blank_paragraphs", { max: 1 });
    expect(doc.texts()).toEqual([
      "start", "", "text",
      "", "",          // the picture paragraph is not blank; one of the two blanks after it stays
      "\f", "",
      "", "", "",      // table cell paragraphs are never removed
      "tail", "",
    ]);
    expect(out.removed).toBe(6);
    expect(doc.deleteLog).toEqual([...doc.deleteLog].sort((a, b) => b - a)); // deleted from the end, so indexes stay valid
    expect(doc.paragraphs.find((p) => p.pictures === 1)).toBeDefined();
    expect(doc.paragraphs.filter((p) => p.tableNestingLevel === 1)).toHaveLength(3);
    expect(doc.texts().filter((t) => t === "\f")).toHaveLength(1);
  });

  it("can remove all blanks with max 0 but never the final paragraph", async () => {
    await data("collapse_blank_paragraphs", { max: 0 });
    expect(doc.texts().filter((t) => t === "")).toHaveLength(1 /* picture */ + 3 /* table */ + 1 /* final */);
    expect(doc.paragraphs.at(-1)!.text).toBe("");
    expect(doc.texts()).toContain("\f");
  });

  it("does not treat a content control paragraph as blank", async () => {
    doc.setParagraphs(["a", "", "", ""]);
    doc.controls.push({ id: 1, tag: "t", title: "", text: "", type: "RichText", para: doc.paragraphs[2]! });
    const out = await data("collapse_blank_paragraphs", { max: 1, preview: true });
    // runs: [1] and [3]; paragraph 2 holds a control, so nothing exceeds max
    expect(out.would_remove).toBe(0);
  });

  it("needs WordApi 1.3 and says so", async () => {
    installWordFake(doc, { sets: { WordApi: "1.2" } });
    expect((await failure("collapse_blank_paragraphs", {})).code).toBe("unsupported");
  });
});
