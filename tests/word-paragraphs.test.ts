import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { executeWordTool } from "../frontend/taskpane/executor/word/index.js";
import { expandParagraphs } from "../frontend/taskpane/executor/word/edit.js";
import { EditDocParagraphsInputSchema } from "../protocol/src/index.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";

let doc: FakeWordDoc;

beforeEach(() => {
  doc = installWordFake(new FakeWordDoc(["Report", "Intro text.", "Middle text.", "Closing text."]));
  doc.paragraphs[0]!.style = "Heading 1";
  doc.paragraphs[0]!.styleBuiltIn = "Heading1";
});
afterEach(() => uninstallWordFake());

const texts = () => doc.paragraphs.map((p) => p.text);

async function run(args: unknown): Promise<any> {
  const result = await executeWordTool("edit_doc_paragraphs", args);
  if (!result.ok) throw new Error(`${result.error.code}: ${result.error.message}`);
  return result.data;
}
async function refused(args: unknown) {
  const result = await executeWordTool("edit_doc_paragraphs", args);
  expect(result.ok).toBe(false);
  return result.ok ? (undefined as never) : result.error;
}

describe("edit_doc_paragraphs: insert", () => {
  it("writes new paragraphs after a paragraph, with the styles asked for", async () => {
    const out = await run({ operation: "insert", paragraph_index: 1, paragraphs: [{ text: "Background", style: "Heading 2" }, "First point.\nSecond point."] });
    expect(texts()).toEqual(["Report", "Intro text.", "Background", "First point.", "Second point.", "Middle text.", "Closing text."]);
    expect(doc.paragraphs[2]!.styleBuiltIn).toBe("Heading2");
    expect(out).toMatchObject({ inserted: 3, first_paragraph_index: 2, last_paragraph_index: 4 });
    expect(out.summary).toContain("after paragraph 1");
  });

  it("makes the line after a heading plain body text instead of another heading", async () => {
    await run({ operation: "insert", paragraph_index: 0, paragraphs: ["Summary line."] });
    expect(doc.paragraphs[1]!.text).toBe("Summary line.");
    expect(doc.paragraphs[1]!.styleBuiltIn).toBe("Normal");
  });

  it("writes at the top with position before, and at the end without an index", async () => {
    await run({ operation: "insert", paragraph_index: 0, position: "before", paragraphs: [{ text: "Draft", style: "Subtitle" }] });
    await run({ operation: "insert", paragraphs: ["Appendix."] });
    expect(texts()[0]).toBe("Draft");
    expect(texts().at(-1)).toBe("Appendix.");
  });

  it("refuses to write inside a table and an index past the end", async () => {
    doc.paragraphs.push(doc.newPara("cell", "Normal", { tableNestingLevel: 1 }));
    expect((await refused({ operation: "insert", paragraph_index: 4, paragraphs: ["x"] })).message).toMatch(/inside a table/);
    expect((await refused({ operation: "insert", paragraph_index: 40, paragraphs: ["x"] })).message).toMatch(/outside the document/);
    expect(texts()).toHaveLength(5);
  });
});

describe("edit_doc_paragraphs: replace and delete", () => {
  it("replaces paragraphs one for one, keeping their style", async () => {
    doc.paragraphs[2]!.style = "Quote";
    const out = await run({ operation: "replace", paragraph_index: 1, count: 2, paragraphs: ["New intro.", "New middle."] });
    expect(texts()).toEqual(["Report", "New intro.", "New middle.", "Closing text."]);
    expect(doc.paragraphs[2]!.style).toBe("Quote");
    expect(out).toMatchObject({ replaced: 2, first_paragraph_index: 1, last_paragraph_index: 2 });
  });

  it("grows and shrinks the passage when the counts differ", async () => {
    await run({ operation: "replace", paragraph_index: 1, count: 1, paragraphs: ["A.", "B.", "C."] });
    expect(texts()).toEqual(["Report", "A.", "B.", "C.", "Middle text.", "Closing text."]);
    const out = await run({ operation: "replace", paragraph_index: 1, count: 4, paragraphs: ["Only one."] });
    expect(texts()).toEqual(["Report", "Only one.", "Closing text."]);
    expect(out).toMatchObject({ replaced: 1, deleted: 3 });
  });

  it("empties a paragraph given empty text and restyles one when asked", async () => {
    await run({ operation: "replace", paragraph_index: 2, paragraphs: [{ text: "", style: "Normal" }] });
    expect(texts()[2]).toBe("");
    await run({ operation: "replace", paragraph_index: 3, paragraphs: [{ text: "Conclusion", style: "heading 2" }] });
    expect(doc.paragraphs[3]!.styleBuiltIn).toBe("Heading2");
  });

  it("deletes paragraphs and refuses a range that runs past the end", async () => {
    const out = await run({ operation: "delete", paragraph_index: 1, count: 2 });
    expect(texts()).toEqual(["Report", "Closing text."]);
    expect(out).toMatchObject({ deleted: 2 });
    expect((await refused({ operation: "delete", paragraph_index: 1, count: 5 })).message).toMatch(/runs past the end/);
  });

  it("says a deletion is only a tracked change when the document tracks changes", async () => {
    doc.changeTrackingMode = "TrackAll";
    const out = await run({ operation: "delete", paragraph_index: 2 });
    expect(out).toMatchObject({ tracked: true });
    expect(out.note).toMatch(/Track Changes is on/);
  });

  it("records the edit as a tracked change when asked", async () => {
    doc.changeTrackingMode = "Off";
    const out = await run({ operation: "delete", paragraph_index: 3, track: true });
    expect(doc.changeTrackingMode).toBe("TrackAll");
    expect(out).toMatchObject({ tracked: true, track_changes_left_on: true });
  });
});

describe("edit_doc_paragraphs arguments", () => {
  it("asks for what each operation needs", () => {
    expect(EditDocParagraphsInputSchema.safeParse({ operation: "insert" }).success).toBe(false);
    expect(EditDocParagraphsInputSchema.safeParse({ operation: "replace", paragraphs: ["x"] }).success).toBe(false);
    expect(EditDocParagraphsInputSchema.safeParse({ operation: "delete", paragraph_index: 1, paragraphs: ["x"] }).success).toBe(false);
    expect(EditDocParagraphsInputSchema.safeParse({ operation: "delete", paragraph_index: 1 }).success).toBe(true);
    expect(EditDocParagraphsInputSchema.safeParse({ operation: "insert", paragraphs: [{ text: "x", font: "Arial" }] }).success).toBe(false);
  });

  it("turns line breaks into paragraphs that share the style", () => {
    expect(expandParagraphs(["a\r\nb", { text: "c\nd", style: "Quote" }])).toEqual([
      { text: "a" }, { text: "b" }, { text: "c", style: "Quote" }, { text: "d", style: "Quote" },
    ]);
  });
});

describe("manage_doc_pictures", () => {
  async function addPicture(after: number, alt = "") {
    const { executeSharedTool } = await import("../frontend/taskpane/host/shared.js");
    const result = await executeSharedTool("word", "insert_image", { source: "x", imageBase64: "AAAA", mimeType: "image/png", sourceLabel: "test", after_paragraph: after, ...(alt ? { alt_text: alt } : {}) });
    if (!result?.ok) throw new Error("insert failed");
  }
  const pictures = (args: unknown) => executeWordTool("manage_doc_pictures", args);

  it("lists pictures in document order with their size, alt text and paragraph", async () => {
    await addPicture(3, "Chart");
    await addPicture(0, "Logo");
    const result = await pictures({ operation: "list" });
    expect(result).toMatchObject({ ok: true, data: { count: 2 } });
    const listed = (result as any).data.pictures;
    expect(listed.map((p: any) => [p.picture_index, p.alt_text, p.width, p.height])).toEqual([[0, "Logo", 200, 100], [1, "Chart", 200, 100]]);
    expect(listed[0]).toHaveProperty("paragraph_text", "");
  });

  it("resizes a picture keeping its proportions and sets its alt text", async () => {
    await addPicture(1);
    const result = await pictures({ operation: "set", picture_index: 0, width: 100, alt_text: "Revenue chart" });
    expect(result).toMatchObject({ ok: true, data: { width: 100, height: 50, alt_text: "Revenue chart" } });
    expect(doc.pictureLog[0]).toMatchObject({ width: 100, height: 50, altTextDescription: "Revenue chart", lockAspectRatio: true });
  });

  it("deletes a picture and refuses an index that does not exist", async () => {
    await addPicture(1);
    expect(await pictures({ operation: "delete", picture_index: 3 })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(texts()).toHaveLength(5);
    expect(await pictures({ operation: "delete", picture_index: 0 })).toMatchObject({ ok: true, data: { remaining: 0, removed_empty_paragraph: true } });
    expect(doc.pictureLog).toHaveLength(0);
    expect(texts()).toEqual(["Report", "Intro text.", "Middle text.", "Closing text."]);
    expect(doc.paragraphs.every((p) => p.pictures === 0)).toBe(true);
  });

  it("keeps the paragraph when it holds text besides the picture", async () => {
    await addPicture(1);
    doc.paragraphs[2]!.text = "Caption text";
    expect(await pictures({ operation: "delete", picture_index: 0 })).toMatchObject({ ok: true, data: { remaining: 0 } });
    expect(texts()).toContain("Caption text");
  });

  it("asks for what set and delete need", async () => {
    expect(await pictures({ operation: "set", picture_index: 0 })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await pictures({ operation: "delete" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });
});
