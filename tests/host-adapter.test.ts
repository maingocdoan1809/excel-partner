import { afterEach, describe, expect, it } from "vitest";
import { detectHost } from "../frontend/taskpane/host/detect.js";
import { activeHost, hostAdapterFor, initHost } from "../frontend/taskpane/host/index.js";
import { HOST_BASELINE, isSupported, requireSupport } from "../frontend/taskpane/host/requirements.js";
import { describeChange } from "../frontend/taskpane/ui/changes.js";
import { PPT_WRITE_TOOLS, PptSelectionSnapshotSchema, WORD_WRITE_TOOLS, WordSelectionSnapshotSchema, exposedToolNamesForHost } from "../protocol/src/index.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";
import { FakeDeck, installFakePowerPoint } from "./support/fake-powerpoint.js";

const g = globalThis as Record<string, any>;
afterEach(() => { delete g.Office; });

function office(host: string | undefined, sets: Record<string, boolean> = {}) {
  g.Office = { context: { ...(host ? { host } : {}), requirements: { isSetSupported: (set: string, version: string) => sets[`${set} ${version}`] ?? sets[set] ?? false } } };
}

describe("detectHost", () => {
  it.each([["Excel", "excel"], ["Word", "word"], ["PowerPoint", "powerpoint"]])("reads Office.context.host=%s", (reported, expected) => {
    office(reported);
    expect(detectHost()).toBe(expected);
  });

  it("falls back to the requirement sets when the host is missing", () => {
    office(undefined, { "ExcelApi": true });
    expect(detectHost()).toBe("excel");
    office(undefined, { "WordApi": true });
    expect(detectHost()).toBe("word");
    office(undefined, { "PowerPointApi": true });
    expect(detectHost()).toBe("powerpoint");
  });

  it("treats a page with no Office behind it as Excel (the browser tab)", () => {
    delete g.Office;
    expect(detectHost()).toBe("excel");
    office(undefined);
    expect(detectHost()).toBe("excel");
  });
});

describe("host adapters", () => {
  it("describe each application", () => {
    expect(["excel", "word", "powerpoint"].map((id) => { const h = hostAdapterFor(id as never); return [h.id, h.app, h.noun, h.hasSheets]; })).toEqual([
      ["excel", "Excel", "workbook", true],
      ["word", "Word", "document", false],
      ["powerpoint", "PowerPoint", "presentation", false],
    ]);
  });

  it("keep write tools per host and never mix them", () => {
    const excel = hostAdapterFor("excel").writeTools, word = hostAdapterFor("word").writeTools, ppt = hostAdapterFor("powerpoint").writeTools;
    expect(excel.has("write_range")).toBe(true);
    expect(excel.has("edit_doc_text")).toBe(false);
    for (const tool of WORD_WRITE_TOOLS) expect(word.has(tool), tool).toBe(true);
    for (const tool of PPT_WRITE_TOOLS) expect(ppt.has(tool), tool).toBe(true);
    expect(word.has("execute_office_js") && ppt.has("execute_office_js")).toBe(true);
    expect(word.has("write_range") || word.has("edit_slide_text") || ppt.has("edit_doc_text") || ppt.has("write_range")).toBe(false);
  });

  it("only offer tools the bridge exposes for that host", () => {
    for (const id of ["word", "powerpoint"] as const) {
      const exposed = new Set(exposedToolNamesForHost(id, true, { playwright: true }));
      for (const tool of hostAdapterFor(id).writeTools) expect(exposed.has(tool), `${id}:${tool}`).toBe(true);
    }
  });

  it("is decided by Office.onReady and not remembered before that", () => {
    office("Word");
    expect(activeHost().id).toBe("word");
    office("PowerPoint");
    expect(activeHost().id).toBe("powerpoint");
    expect(initHost("Excel").id).toBe("excel");
    office("Word");
    expect(activeHost().id).toBe("excel");
    initHost("Word");
    expect(activeHost().id).toBe("word");
    office("PowerPoint");
    expect(initHost(undefined).id).toBe("powerpoint"); // nothing reported: worked out from Office, then remembered
    office("Word");
    expect(activeHost().id).toBe("powerpoint");
  });
});

describe("Word adapter", () => {
  afterEach(() => uninstallWordFake());

  function wordPane() {
    const doc = new FakeWordDoc();
    doc.setParagraphs([["Heading", "Heading 1"], "Some text to select."]);
    installWordFake(doc);
    g.Office.context.host = "Word";
    return doc;
  }

  it("answers get_selection with a Word snapshot", async () => {
    wordPane();
    const result = await hostAdapterFor("word").execute("get_selection", {});
    expect(result.ok).toBe(true);
    if (result.ok) expect(WordSelectionSnapshotSchema.safeParse(result.data).success).toBe(true);
  });

  it("refuses Excel and PowerPoint tools with unknown_tool", async () => {
    wordPane();
    for (const tool of ["read_range", "write_range", "edit_slide_text", "get_sheets"]) {
      expect(await hostAdapterFor("word").execute(tool, {}), tool).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
    }
  });

  it("runs execute_office_js inside Word.run with Word in scope", async () => {
    wordPane();
    const result = await hostAdapterFor("word").execute("execute_office_js", { code: "return { ns: typeof Word.run, ctx: typeof context.sync, office: typeof Office };" });
    expect(result).toEqual({ ok: true, data: { result: { ns: "function", ctx: "function", office: "object" } } });
  });

  it("reports code that does not compile and results that are too large", async () => {
    wordPane();
    expect(await hostAdapterFor("word").execute("execute_office_js", { code: "return (" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await hostAdapterFor("word").execute("execute_office_js", { code: "return 'x'.repeat(40000);" })).toMatchObject({ ok: false, error: { code: "size_limit_exceeded" } });
  });

  it("acknowledges create_artifact", async () => {
    wordPane();
    expect(await hostAdapterFor("word").execute("create_artifact", { filename: "a.md", content: "hello" })).toEqual({ ok: true, data: { filename: "a.md", characters: 5, shown: true } });
  });

  it("names the open document from its URL", async () => {
    const doc = wordPane();
    doc.url = "https://contoso.sharepoint.com/sites/x/Shared%20Documents/Plan%20Q3.docx";
    expect(await hostAdapterFor("word").readName()).toBe("Plan Q3.docx");
    doc.url = "";
    expect(await hostAdapterFor("word").readName()).toBeNull();
  });
});

describe("PowerPoint adapter", () => {
  function pptPane() {
    const installed = installFakePowerPoint(new FakeDeck());
    g.Office.context.host = "PowerPoint";
    return installed;
  }

  it("answers get_selection with a PowerPoint snapshot and runs PowerPoint.run for execute_office_js", async () => {
    const installed = pptPane();
    try {
      const selection = await hostAdapterFor("powerpoint").execute("get_selection", {});
      expect(selection.ok).toBe(true);
      if (selection.ok) expect(PptSelectionSnapshotSchema.safeParse(selection.data).success).toBe(true);
      expect(await hostAdapterFor("powerpoint").execute("execute_office_js", { code: "return typeof PowerPoint.run;" })).toEqual({ ok: true, data: { result: "function" } });
      expect(await hostAdapterFor("powerpoint").execute("read_doc_section", {})).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
    } finally { installed.restore(); }
  });
});

describe("requirement gate", () => {
  it("passes only when every set is present, and never throws without Office", () => {
    office("Word", { "WordApi 1.4": true });
    expect(isSupported(["WordApi", "1.4"])).toBe(true);
    expect(isSupported(["WordApi", "1.4"], ["WordApi", "1.6"])).toBe(false);
    delete g.Office;
    expect(isSupported(["WordApi", "1.1"])).toBe(false);
  });

  it("stops with unsupported and an actionable hint", () => {
    office("Word", {});
    expect(() => requireSupport("Comments", ["WordApi", "1.4"], "Update Word.")).toThrowError(/Comments needs WordApi 1\.4.*Update Word\./);
    expect(HOST_BASELINE).toEqual({ excel: ["ExcelApi", "1.9"], word: ["WordApi", "1.1"], powerpoint: ["PowerPointApi", "1.2"] });
  });
});

describe("change list for Word and PowerPoint", () => {
  it("shows the tool's own summary", () => {
    const writeTools = hostAdapterFor("word").writeTools;
    expect(describeChange("edit_doc_text", {}, { ok: true, data: { summary: "Replaced 1 match" } }, writeTools)).toMatchObject({ ok: true, summary: "Replaced 1 match" });
    expect(describeChange("edit_doc_text", {}, { ok: true, data: {} }, writeTools)).toMatchObject({ summary: "edit_doc_text" });
    expect(describeChange("read_doc_section", {}, { ok: true, data: {} }, writeTools)).toBeNull();
  });

  it("does not list calls that only read", () => {
    const writeTools = new Set(["edit_table", "manage_content_controls", "manage_headers_footers", "collapse_blank_paragraphs", "edit_slide_table"]);
    const okResult = { ok: true as const, data: {} };
    expect(describeChange("edit_table", { operation: "read" }, okResult, writeTools)).toBeNull();
    expect(describeChange("edit_slide_table", { operation: "read" }, okResult, writeTools)).toBeNull();
    expect(describeChange("manage_content_controls", { operation: "list" }, okResult, writeTools)).toBeNull();
    expect(describeChange("manage_headers_footers", { operation: "read" }, okResult, writeTools)).toBeNull();
    expect(describeChange("collapse_blank_paragraphs", { preview: true }, okResult, writeTools)).toBeNull();
    expect(describeChange("edit_table", { operation: "add_rows" }, okResult, writeTools)).not.toBeNull();
  });

});
