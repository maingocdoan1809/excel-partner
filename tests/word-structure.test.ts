import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { executeWordTool } from "../frontend/taskpane/executor/word/index.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";

let doc: FakeWordDoc;

beforeEach(() => {
  doc = installWordFake(new FakeWordDoc(["Title line", "First paragraph about revenue.", "Second paragraph about costs."]));
  doc.paragraphs[0]!.style = "Heading 1";
  doc.paragraphs[0]!.styleBuiltIn = "Heading1";
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

describe("apply_style", () => {
  it("styles paragraphs by index with a built-in id and formatting", async () => {
    const out = await data("apply_style", { target: "paragraphs", paragraph_index: 1, count: 2, style: "Heading 2", font: { bold: true, size: 14, color: "#112233", highlight: null }, alignment: "Centered", space_after: 6 });
    expect(out).toMatchObject({ paragraphs: 2, target: "paragraphs" });
    const [, a, b] = doc.paragraphs;
    expect([a!.styleBuiltIn, b!.styleBuiltIn]).toEqual(["Heading2", "Heading2"]);
    expect(a!.font).toMatchObject({ bold: true, size: 14, color: "#112233", highlightColor: null });
    expect([a!.alignment, a!.spaceAfter]).toEqual(["Centered", 6]);
    expect(out.summary).toContain("2 paragraphs");
  });

  it("uses the style name for a custom style and when styleBuiltIn is not available", async () => {
    await data("apply_style", { target: "paragraphs", paragraph_index: 1, style: "My Custom Style" });
    expect(doc.paragraphs[1]!.style).toBe("My Custom Style");
    installWordFake(doc, { sets: { WordApi: "1.2" } });
    await data("apply_style", { target: "paragraphs", paragraph_index: 2, style: "Quote" });
    expect(doc.paragraphs[2]!.style).toBe("Quote");
  });

  it("formats the selection, and refuses a font change on a bare cursor", async () => {
    doc.select(1, 0, 5);
    await data("apply_style", { font: { italic: true } });
    expect(doc.paragraphs[1]!.font.italic).toBe(true);
    doc.select(2, 3);
    const error = await failure("apply_style", { font: { italic: true } });
    expect(error.code).toBe("invalid_arguments");
    expect((await data("apply_style", { style: "Quote" })).paragraphs).toBe(1);
  });

  it("rejects an index outside the document and a call with nothing to apply", async () => {
    expect((await failure("apply_style", { target: "paragraphs", paragraph_index: 40, style: "Normal" })).code).toBe("invalid_arguments");
    expect((await failure("apply_style", {})).code).toBe("invalid_arguments");
  });
});

describe("tables", () => {
  it("inserts a table at the end and returns its index", async () => {
    const out = await data("insert_table", { values: [["Name", "Qty"], ["Apple", "3"]], style: "Grid Table 4 - Accent 1" });
    expect(out).toMatchObject({ rows: 2, columns: 2, table_index: 0, method: "table_api" });
    expect(doc.tables[0]).toMatchObject({ rows: [["Name", "Qty"], ["Apple", "3"]], headerRowCount: 1, style: "Grid Table 4 - Accent 1" });
    expect((await data("insert_table", { values: [["x"]], header_row: false })).table_index).toBe(1);
    expect(doc.tables[1]!.headerRowCount).toBe(0);
  });

  it("finds the index of a table inserted after a paragraph", async () => {
    await data("insert_table", { values: [["end"]] });
    const out = await data("insert_table", { values: [["A", "B"]], after_paragraph: 0 });
    expect(out.table_index).toBe(0);
    expect(doc.orderedTables()[0]!.rows[0]).toEqual(["A", "B"]);
  });

  it("keeps the table when its style is wrong and says so", async () => {
    const out = await data("insert_table", { values: [["a"]], style: "Bad style" });
    expect(out.style_note).toContain("could not be applied");
    expect(doc.tables).toHaveLength(1);
  });

  it("falls back to an HTML table below WordApi 1.3", async () => {
    installWordFake(doc, { sets: { WordApi: "1.2" } });
    const out = await data("insert_table", { values: [["a<b", "c"]] });
    expect(out).toMatchObject({ method: "html", table_index: null });
    expect(doc.htmlInserts[0]!.html).toContain("<th>a&lt;b</th>");
  });

  describe("edit_table", () => {
    beforeEach(async () => {
      await data("insert_table", { values: [["h1", "h2"], ["a", "b"], ["c", "d"]] });
    });

    it("reads values", async () => {
      expect(await data("edit_table", { table_index: 0, operation: "read" })).toMatchObject({ rows: 3, columns: 2, header_rows: 1, values: [["h1", "h2"], ["a", "b"], ["c", "d"]] });
    });

    it("sets a cell, adds and deletes rows and columns, sets the style and deletes the table", async () => {
      await data("edit_table", { table_index: 0, operation: "set_cell", row: 1, column: 1, text: "B!" });
      expect(doc.tables[0]!.rows[1]).toEqual(["a", "B!"]);
      expect(await data("edit_table", { table_index: 0, operation: "add_rows", row: 0, count: 2, position: "after" })).toMatchObject({ rows: 5, columns: 2 });
      expect(await data("edit_table", { table_index: 0, operation: "delete_rows", row: 1, count: 2 })).toMatchObject({ rows: 3 });
      expect(await data("edit_table", { table_index: 0, operation: "add_columns", column: 0, position: "before" })).toMatchObject({ columns: 3 });
      expect(await data("edit_table", { table_index: 0, operation: "delete_columns", column: 0 })).toMatchObject({ columns: 2 });
      await data("edit_table", { table_index: 0, operation: "set_style", style: "Plain Table 1" });
      expect(doc.tables[0]!.style).toBe("Plain Table 1");
      await data("edit_table", { table_index: 0, operation: "merge", row: 0, column: 0, to_row: 0, to_column: 1 });
      expect(doc.tables[0]!.merges).toEqual([[0, 0, 0, 1]]);
      expect(await data("edit_table", { table_index: 0, operation: "delete_table" })).toMatchObject({ deleted: true });
      expect(doc.tables).toHaveLength(0);
    });

    it("validates bounds and keeps at least one row and column", async () => {
      expect((await failure("edit_table", { table_index: 5, operation: "read" })).code).toBe("range_not_found");
      expect((await failure("edit_table", { table_index: 0, operation: "set_cell", row: 9, column: 0, text: "x" })).code).toBe("invalid_arguments");
      expect((await failure("edit_table", { table_index: 0, operation: "delete_rows", row: 0, count: 3 })).message).toContain("delete_table");
      expect((await failure("edit_table", { table_index: 0, operation: "delete_columns", column: 0, count: 2 })).code).toBe("invalid_arguments");
    });

    it("needs WordApi 1.3, and 1.4 for merge", async () => {
      installWordFake(doc, { sets: { WordApi: "1.3" } });
      expect((await failure("edit_table", { table_index: 0, operation: "merge", row: 0, column: 0, to_row: 1, to_column: 1 })).message).toContain("WordApi 1.4");
      installWordFake(doc, { sets: { WordApi: "1.2" } });
      expect((await failure("edit_table", { table_index: 0, operation: "read" })).code).toBe("unsupported");
    });

    it("cuts a huge table read to 30,000 characters", async () => {
      doc.tables[0]!.rows = Array.from({ length: 300 }, (_, r) => [`row ${r}`, "y".repeat(400)]);
      const out = await data("edit_table", { table_index: 0, operation: "read" });
      expect(JSON.stringify(out).length).toBeLessThanOrEqual(30_000);
      expect(out.truncated).toBe(true);
      expect(out.rows).toBe(300);
    });
  });
});

describe("content controls", () => {
  it("inserts, lists, fills and deletes controls", async () => {
    doc.select(1, 0, 5);
    const inserted = await data("manage_content_controls", { operation: "insert", tag: "client", title: "Client name", text: "ACME" });
    expect(inserted).toMatchObject({ tag: "client", title: "Client name" });
    const listed = await data("manage_content_controls", { operation: "list" });
    expect(listed.controls).toEqual([expect.objectContaining({ tag: "client", text: "ACME" })]);
    await data("manage_content_controls", { operation: "set_text", tag: "client", text: "Globex" });
    expect(doc.controls[0]!.text).toBe("Globex");
    await data("manage_content_controls", { operation: "set_text", title: "Client name", text: "Initech" });
    expect(doc.controls[0]!.text).toBe("Initech");
    expect(await data("manage_content_controls", { operation: "delete", tag: "client" })).toMatchObject({ deleted: 1 });
    expect(doc.controls).toHaveLength(0);
    expect((await failure("manage_content_controls", { operation: "set_text", tag: "gone", text: "x" })).code).toBe("range_not_found");
  });
});

describe("fields", () => {
  it("inserts a field at the selection or the end and refreshes it", async () => {
    const out = await data("insert_field", { field_type: "Page" });
    expect(out).toMatchObject({ field_type: "Page", code: "Page", result: "‹Page›" });
    expect(doc.fields[0]!.updates).toBe(1);
    await data("insert_field", { field_type: "Ref", text: "chapter1", location: "end" });
    expect(doc.fields[1]!.code).toBe("Ref chapter1");
  });

  it("asks for the text a Ref or Seq field needs", async () => {
    expect((await failure("insert_field", { field_type: "Ref" })).message).toContain("bookmark name");
    expect((await failure("insert_field", { field_type: "Seq" })).code).toBe("invalid_arguments");
  });

  it("updates every field and counts the ones that cannot be refreshed", async () => {
    await data("insert_field", { field_type: "Page" });
    await data("insert_field", { field_type: "NumPages" });
    expect(await data("update_fields")).toMatchObject({ fields: 2, updated: 2, failed: 0 });
    doc.fields.push({ type: "Bad", code: "Bad", result: "", updates: 0 });
    expect(await data("update_fields")).toMatchObject({ fields: 3, updated: 2, failed: 1 });
    doc.fields.length = 0;
    expect((await data("update_fields")).summary).toContain("no fields");
  });

  it("is unsupported below WordApi 1.5", async () => {
    installWordFake(doc, { sets: { WordApi: "1.4" } });
    expect((await failure("insert_field", { field_type: "Page" })).code).toBe("unsupported");
    expect((await failure("update_fields")).message).toContain("WordApi 1.5");
    expect((await failure("insert_footnote", { anchor_text: "revenue", text: "x" })).code).toBe("unsupported");
  });
});

describe("table of contents", () => {
  it("inserts and updates a TOC on desktop", async () => {
    const out = await data("insert_toc", { levels: 2, after_paragraph: 0 });
    expect(out).toMatchObject({ levels: 2, tables_of_contents: 1 });
    expect(doc.tocs[0]!.options).toMatchObject({ upperHeadingLevel: 1, lowerHeadingLevel: 2 });
    expect(doc.paragraphs).toHaveLength(4);
    expect(await data("update_toc")).toMatchObject({ updated: 1, entries_refreshed: 1 });
  });

  it("only refreshes page numbers when the field API is missing, and tells the model", async () => {
    await data("insert_toc", {});
    installWordFake(doc, { sets: { WordApi: "1.3", WordApiDesktop: "1.4" } });
    const out = await data("update_toc");
    expect(out.entries_refreshed).toBe(0);
    expect(out.note).toContain("page numbers");
    expect(doc.tocs[0]!.updatedPageNumbers).toBe(1);
  });

  it("fails with a clear unsupported error on Word on the web", async () => {
    installWordFake(doc, { sets: { WordApi: "1.6" } });
    for (const tool of ["insert_toc", "update_toc"]) {
      const error = await failure(tool, {});
      expect(error.code).toBe("unsupported");
      expect(error.message).toContain("WordApiDesktop 1.4");
      expect(error.message).toContain("Word on the web");
    }
  });

  it("says there is no TOC to update", async () => {
    expect((await failure("update_toc")).code).toBe("range_not_found");
  });
});

describe("hyperlinks, bookmarks, footnotes", () => {
  it("turns anchor text into a link, or inserts linked text at the selection", async () => {
    await data("insert_hyperlink", { text: "ignored", url: "https://example.com", anchor_text: "revenue" });
    expect(doc.hyperlinks).toEqual([{ text: "revenue", url: "https://example.com" }]);
    await data("insert_hyperlink", { text: "Docs", url: "mailto:a@b.c" });
    expect(doc.hyperlinks[1]).toEqual({ text: "Docs", url: "mailto:a@b.c" });
  });

  it("uses HTML below WordApi 1.3", async () => {
    installWordFake(doc, { sets: { WordApi: "1.2" } });
    const out = await data("insert_hyperlink", { text: "x", url: "https://e.com/?a=1&b=2", anchor_text: "revenue" });
    expect(out.method).toBe("html");
    expect(doc.htmlInserts[0]!.html).toBe('<a href="https://e.com/?a=1&amp;b=2">revenue</a>');
  });

  it("adds a bookmark and warns when it moves an existing one", async () => {
    const first = await data("insert_bookmark", { name: "Costs_1", anchor_text: "costs" });
    expect(first.replaced_existing).toBeUndefined();
    expect(doc.bookmarks.get("Costs_1")).toBe("costs");
    const again = await data("insert_bookmark", { name: "costs_1", anchor_text: "revenue" });
    expect(again.replaced_existing).toBe(true);
  });

  it("adds a footnote or endnote after the anchor", async () => {
    await data("insert_footnote", { anchor_text: "revenue", text: "Source: ledger" });
    await data("insert_footnote", { anchor_text: "costs", text: "See annex", kind: "endnote" });
    expect(doc.notes).toEqual([{ kind: "footnote", text: "Source: ledger", after: "revenue" }, { kind: "endnote", text: "See annex", after: "costs" }]);
  });

  it("bookmark is unsupported below WordApi 1.4; footnote anchor must exist", async () => {
    expect((await failure("insert_footnote", { anchor_text: "absent", text: "x" })).code).toBe("range_not_found");
    installWordFake(doc, { sets: { WordApi: "1.3" } });
    expect((await failure("insert_bookmark", { name: "x" })).code).toBe("unsupported");
  });
});

describe("headers, footers, page layout", () => {
  it("sets and reads a footer and warns about first-page headers", async () => {
    await data("manage_headers_footers", { operation: "set", part: "footer", text: "Confidential" });
    expect(doc.sections[0]!.footer.Primary).toBe("Confidential");
    expect(await data("manage_headers_footers", { operation: "read", part: "footer" })).toMatchObject({ text: "Confidential", sections: 1 });
    const first = await data("manage_headers_footers", { operation: "set", type: "FirstPage", text: "Cover" });
    expect(first.note).toContain("Different first page");
    expect((await failure("manage_headers_footers", { operation: "read", section: 3 })).code).toBe("invalid_arguments");
  });

  it("sets orientation, paper and margins on desktop and reports the result", async () => {
    const out = await data("set_page_layout", { orientation: "Landscape", paper: "A4", margins: { top: 36, left: 54 } });
    expect(out).toMatchObject({ orientation: "Landscape", paper: "A4", margins: { top: 36, bottom: 72, left: 54, right: 72 }, page: { width: 792, height: 612 } });
  });

  it("fails as unsupported on Word on the web", async () => {
    installWordFake(doc, { sets: { WordApi: "1.6" } });
    const error = await failure("set_page_layout", { orientation: "Landscape" });
    expect(error.code).toBe("unsupported");
    expect(error.message).toContain("WordApiDesktop 1.3");
  });
});

describe("results", () => {
  it("every write result carries a short summary", async () => {
    const results = [
      await data("edit_doc_text", { old_text: "revenue", new_text: "sales" }),
      await data("edit_doc_list", { operation: "bullets", items: ["a"] }),
      await data("collapse_blank_paragraphs", {}),
      await data("add_comment", { anchor_text: "costs", text: "c" }),
      await data("set_track_changes", { enabled: true }),
      await data("apply_style", { target: "paragraphs", paragraph_index: 0, style: "Normal" }),
      await data("insert_table", { values: [["a"]] }),
      await data("edit_table", { table_index: 0, operation: "set_cell", row: 0, column: 0, text: "x" }),
      await data("insert_field", { field_type: "Page" }),
      await data("update_fields"),
      await data("insert_toc", {}),
      await data("update_toc"),
      await data("insert_hyperlink", { text: "t", url: "https://e.com" }),
      await data("insert_bookmark", { name: "b", anchor_text: "costs" }),
      await data("manage_headers_footers", { operation: "set", text: "h" }),
      await data("insert_footnote", { anchor_text: "costs", text: "n" }),
      await data("set_page_layout", { orientation: "Landscape" }),
    ];
    for (const result of results) {
      expect(typeof result.summary).toBe("string");
      expect(result.summary.length).toBeGreaterThan(0);
      expect(result.summary.length).toBeLessThan(200);
    }
  });

  it("refuses to return more than 30,000 characters from any tool", async () => {
    doc.setParagraphs(Array.from({ length: 4000 }, (_, i) => [`Heading number ${i} ${"h".repeat(150)}`, "Heading 1"] as [string, string]));
    const out = await data("get_doc_outline", { max_headings: 1000 });
    expect(JSON.stringify(out).length).toBeLessThanOrEqual(30_000);
    expect(out.truncated).toBe(true);
    expect(out.headingCount).toBe(4000);
  });
});
