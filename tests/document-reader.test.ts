import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocumentReadError, documentFormatOf, readDocumentFile } from "../bridge/src/document-reader.js";
import { buildDocx, buildPptx, para, sampleDeck, sampleDocx } from "./support/office-files.js";

const read = (name: string, bytes: Buffer, options: Partial<{ from: number; count: number; maxChars: number }> = {}): any =>
  readDocumentFile(name, bytes, { from: 1, maxChars: 20_000, ...options });

describe("which files the reader takes", () => {
  it("goes by extension", () => {
    expect(documentFormatOf("Report.DOCX")).toBe("docx");
    expect(documentFormatOf("deck.pptx")).toBe("pptx");
    expect(documentFormatOf("template.potx")).toBe("pptx");
    expect(documentFormatOf("book.xlsx")).toBeNull();
    expect(documentFormatOf("old.doc")).toBeNull();
  });

  it("refuses other types, non-zip bytes and zips that are not Office files, each with a message that says what to do", () => {
    expect(() => read("book.xlsx", sampleDocx())).toThrow(/read_workbook_file/);
    expect(() => read("old.docx", Buffer.from("not a zip at all"))).toThrow(/not a valid Office document/);
    const empty = Buffer.from(zipSync({ "hello.txt": strToU8("hi") }));
    expect(() => read("fake.docx", empty)).toThrow(/not a Word document/);
    expect(() => read("fake.pptx", empty)).toThrow(/not a PowerPoint file/);
    expect(() => read("book.xlsx", sampleDocx())).toThrow(DocumentReadError);
  });

  it("will not inflate a part that claims to be huge", () => {
    const bomb = Buffer.from(zipSync({ "word/document.xml": new Uint8Array(13 * 1024 * 1024) }));
    expect(() => read("bomb.docx", bomb)).toThrow(/too large/);
  });
});

describe("Word files", () => {
  const result = read("report.docx", sampleDocx());

  it("returns blocks in reading order with styles, heading levels and list marks", () => {
    expect(result.format).toBe("docx");
    expect(result.blocks[0]).toMatchObject({ index: 1, style: "heading 1", level: 1, text: "Annual report" });
    expect(result.blocks[1]).toMatchObject({ index: 2, level: 2, text: "Overview" });
    expect(result.blocks.find((b: any) => b.list)).toMatchObject({ text: "First point" });
    expect(result.blocks.at(-1)).toMatchObject({ text: "Closing remarks." });
  });

  it("reads the accepted wording of tracked changes and shows a text box once", () => {
    const texts = result.blocks.map((b: any) => b.text);
    expect(texts).toContain("Revenue grew steadily");
    expect(texts.join("|")).not.toContain("sharply");
    expect(texts.find((t: string) => t?.startsWith("Before box."))).toBe("Before box.Boxed words");
  });

  it("returns a table as one block with rows and decoded entities", () => {
    const table = result.blocks.find((b: any) => b.table);
    expect(table.table).toEqual({ rowCount: 2, rows: [["Region", "Sales"], ["EMEA", "10 & up"]] });
  });

  it("gives an outline of the headings and the comments on the first page", () => {
    expect(result.outline).toEqual([{ index: 1, level: 1, text: "Annual report" }, { index: 2, level: 2, text: "Overview" }]);
    expect(result.comments).toEqual([{ author: "Lan", text: "Check this number" }]);
    expect(result.blockCount).toBe(result.blocks.length);
  });

  it("pages by position and by character budget, and says where to continue", () => {
    const body = Array.from({ length: 50 }, (_, i) => para(`Paragraph number ${i + 1} ${"x".repeat(100)}`)).join("");
    const doc = buildDocx(body);
    const first = read("long.docx", doc, { maxChars: 1500 });
    expect(first.truncated).toBe(true);
    expect(first.returned).toBeGreaterThan(3);
    expect(first.returned).toBeLessThan(20);
    expect(JSON.stringify(first.blocks).length).toBeLessThan(1700);
    const second = read("long.docx", doc, { from: first.nextFrom, count: 2 });
    expect(second.blocks.map((b: any) => b.index)).toEqual([first.nextFrom, first.nextFrom + 1]);
    expect(second.outline).toBeUndefined();
    expect(read("long.docx", doc, { from: 50, count: 10 })).toMatchObject({ returned: 1, truncated: false });
  });
});

describe("PowerPoint files", () => {
  const result = read("deck.pptx", sampleDeck());

  it("lists slides in presentation order even when the files are numbered differently", () => {
    expect(result.slideCount).toBe(3);
    expect(result.slides.map((s: any) => s.title)).toEqual(["Quarterly Review", "Numbers", "Backup"]);
    expect(result.slides.map((s: any) => s.index)).toEqual([1, 2, 3]);
  });

  it("reads shapes, tables, picture alt text and speaker notes, and leaves out footers and slide numbers", () => {
    const [one, two, three] = result.slides;
    expect(one.notes).toBe("Say thank you first.");
    expect(one.shapes.map((s: any) => s.name)).toEqual(["Title 1", "Content 2"]);
    expect(JSON.stringify(one)).not.toContain("Confidential");
    expect(two.shapes.find((s: any) => s.type === "table")).toMatchObject({ rowCount: 2, table: [["Region", "Sales"], ["EMEA", "10"]] });
    expect(two.shapes.find((s: any) => s.type === "picture")).toMatchObject({ alt: "Company logo" });
    expect(three.hidden).toBe(true);
    expect(three.shapes.find((s: any) => s.name === "Content 2").text).toBe("Hidden detail\n  nested");
  });

  it("returns only the slides asked for and where to continue", () => {
    const window = read("deck.pptx", sampleDeck(), { from: 2, count: 1 });
    expect(window.slides.map((s: any) => s.index)).toEqual([2]);
    expect(window).toMatchObject({ truncated: true, nextFrom: 3 });
    const budget = read("deck.pptx", sampleDeck(), { maxChars: 500 });
    expect(budget.slides.length).toBeGreaterThanOrEqual(1);
    expect(budget.slides.length).toBeLessThan(3);
    expect(budget.nextFrom).toBe(budget.slides.length + 1);
  });

  it("copes with a slide that has no text and a deck of one slide", () => {
    const deck = buildPptx([{}]);
    expect(read("empty.pptx", deck)).toMatchObject({ slideCount: 1, slides: [{ index: 1, shapes: [] }] });
  });
});
