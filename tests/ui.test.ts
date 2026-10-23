import { describe, expect, it } from "vitest";
import { WRITE_TOOLS } from "../addin/taskpane/executor/index.js";
import { describeChange } from "../addin/taskpane/ui/changes.js";
import { renderMarkdown } from "../addin/taskpane/ui/markdown.js";
import { filterCommands, leadingCommand, slashBadgeRange, slashQuery } from "../addin/taskpane/ui/slash.js";
import { activeFileMention, atomicBadgeDeletion, atomicFileMentionDeletion, fileMentionRanges } from "../addin/taskpane/ui/file-mentions.js";

describe("file mentions", () => {
  it("opens a fresh query for a second @ on the same line", () => {
    const names = ["first file.md"];
    expect(activeFileMention("Check @first file.md", names)).toBeNull();
    expect(activeFileMention("Check @first file.md ", names)).toBeNull();
    expect(activeFileMention("Check @first file.md @", names)).toEqual({ start: 21, query: "" });
    expect(activeFileMention("Check @first file.md @sec", names)).toEqual({ start: 21, query: "sec" });
    expect(activeFileMention("email@example.com", names)).toBeNull();
  });

  it("marks only selected file references as badges", () => {
    expect(fileMentionRanges("Compare @first file.md with @second.zip", ["first file.md", "second.zip"]))
      .toEqual([{ start: 8, end: 22, name: "first file.md" }, { start: 28, end: 39, name: "second.zip" }]);
    expect(fileMentionRanges("Try @first file.md2", ["first file.md"])).toEqual([]);
  });

  it("deletes a file mention as one unit", () => {
    const text = "Compare @first.md @second.zip please";
    const names = ["first.md", "second.zip"];
    expect(atomicFileMentionDeletion(text, 18, 18, "deleteContentBackward", names)).toEqual({ start: 8, end: 18 });
    expect(atomicFileMentionDeletion(text, 18, 18, "deleteContentForward", names)).toEqual({ start: 18, end: 29 });
    expect(atomicFileMentionDeletion(text, 12, 14, "deleteContentBackward", names)).toEqual({ start: 8, end: 18 });
    expect(atomicFileMentionDeletion(text, 8, 8, "deleteContentBackward", names)).toBeNull();
    expect(atomicFileMentionDeletion(text, 31, 31, "deleteContentBackward", names)).toBeNull();
  });

  it("deletes a chosen slash skill as one unit, even with a file mention selected", () => {
    const text = "/sample-bd-workbook use @notes.md";
    const skill = slashBadgeRange(text, "/sample-bd-workbook");
    const badges = [skill!, ...fileMentionRanges(text, ["notes.md"])];
    expect(skill).toEqual({ start: 0, end: 19 });
    expect(atomicBadgeDeletion(text, 20, 20, "deleteContentBackward", badges)).toEqual({ start: 0, end: 20 });
    expect(atomicBadgeDeletion(text, 5, 9, "deleteContentBackward", badges)).toEqual({ start: 0, end: 20 });
    expect(atomicBadgeDeletion(text, 0, 0, "deleteContentForward", badges)).toEqual({ start: 0, end: 20 });
    expect(atomicBadgeDeletion(text, 22, 22, "deleteContentBackward", badges)).toBeNull();
  });
});

describe("renderMarkdown", () => {
  it("escapes HTML before anything else", () => {
    expect(renderMarkdown('<img src=x onerror="alert(1)"> & **hi**')).toBe(
      "<p>&lt;img src=x onerror=\"alert(1)\"&gt; &amp; <strong>hi</strong></p>",
    );
  });

  it("renders inline code, emphasis, lists and code fences", () => {
    const html = renderMarkdown("Use `SUM`\n- one\n- *two*\n1. a\n```\nx < 1\n```");
    expect(html).toBe(
      "<p>Use <code>SUM</code></p><ul><li>one</li><li><em>two</em></li></ul><ol><li>a</li></ol><pre><code>x &lt; 1</code></pre>",
    );
  });

  it("renders the reported AP News Markdown link as a clickable external link", () => {
    expect(renderMarkdown("Nhấp vào [liên kết AP News](https://apnews.com/article/38d37c4c18bf8355b02d85b5a73ed642)."))
      .toBe('<p>Nhấp vào <a href="https://apnews.com/article/38d37c4c18bf8355b02d85b5a73ed642" target="_blank" rel="noopener noreferrer">liên kết AP News</a>.</p>');
  });

  it("keeps links inside code literal, handles URL parentheses and escapes attributes", () => {
    const html = renderMarkdown('`[code](https://example.com)` [**source**](https://example.com/a_(b)?x=1&y=2)');
    expect(html).toContain('<code>[code](https://example.com)</code>');
    expect(html).toContain('href="https://example.com/a_(b)?x=1&amp;y=2"');
    expect(html).toContain('<strong>source</strong></a>');
    expect(renderMarkdown('[quote](https://example.com/"quoted")')).toContain('href="https://example.com/&quot;quoted&quot;"');
    expect(renderMarkdown('```\n[x](https://example.com)\n```')).not.toContain('<a ');
  });

  it("leaves unsafe schemes and incomplete streamed links as text", () => {
    for (const text of ['[x](javascript:alert(1))', '[x](data:text/html,test)', '[x](//example.com)', '[x](https://example.com']) {
      expect(renderMarkdown(text)).not.toContain('<a ');
    }
  });

  it("copes with an unterminated code fence while streaming", () => {
    expect(renderMarkdown("```\nfoo")).toBe("<pre><code>foo</code></pre>");
  });
});

describe("describeChange", () => {
  const ok = (data: unknown) => ({ ok: true as const, data });

  it("ignores read-only tools", () => {
    expect(describeChange("read_range", {}, ok({}), WRITE_TOOLS)).toBeNull();
    expect(describeChange("list_shapes", {}, ok({}), WRITE_TOOLS)).toBeNull();
  });

  it("describes writes with a jump target", () => {
    expect(
      describeChange("write_range", {}, ok({ sheet: "S", address: "A1:B2", rowsWritten: 2, columnsWritten: 2 }), WRITE_TOOLS),
    ).toEqual({ tool: "write_range", ok: true, summary: "Wrote 2×2 at S!A1:B2", target: { sheet: "S", address: "A1:B2" } });
  });

  it("describes shapes and connectors", () => {
    expect(describeChange("add_shape", {}, ok({ name: "a", type: "Rectangle" }), WRITE_TOOLS)?.summary).toBe('Added Rectangle "a"');
    expect(describeChange("add_connector", {}, ok({ from: "a", to: "b" }), WRITE_TOOLS)?.summary).toBe('Connected "a" → "b"');
  });

  it("records failures, including refused overwrites", () => {
    const entry = describeChange(
      "write_range",
      {},
      { ok: false, error: { code: "would_overwrite", message: "Would overwrite 1 non-empty cell: A1." } },
      WRITE_TOOLS,
    );
    expect(entry).toMatchObject({ ok: false, summary: "write_range failed: Would overwrite 1 non-empty cell: A1.", target: null });
  });
});

describe("slash menu helpers", () => {
  const commands = [
    { name: "sample-bd-workbook", description: "Quy tắc Excel BD của công ty", kind: "skill" as const },
    { name: "excel-report", description: "Báo cáo Excel kiểu Nhật", kind: "skill" as const },
    { name: "clear", description: "Clear this conversation", kind: "builtin" as const },
  ];

  it("recognises the command word only while it is being typed", () => {
    expect(slashQuery("/")).toBe("");
    expect(slashQuery("/Ast")).toBe("ast");
    expect(slashQuery("/sample-bd-workbook ")).toBeNull();
    expect(slashQuery("hello /x")).toBeNull();
    expect(slashQuery("/a b")).toBeNull();
  });

  it("ranks prefix matches before substring and description matches", () => {
    expect(filterCommands(commands, "").map((c) => c.name)).toEqual(["sample-bd-workbook", "excel-report", "clear"]);
    expect(filterCommands(commands, "excel").map((c) => c.name)).toEqual(["excel-report", "sample-bd-workbook"]);
    expect(filterCommands(commands, "nhật").map((c) => c.name)).toEqual(["excel-report"]);
    expect(filterCommands(commands, "zzz")).toEqual([]);
  });

  it("splits a leading command for highlighting", () => {
    expect(leadingCommand("/sample-bd-workbook vẽ flow")).toEqual({ command: "/sample-bd-workbook", rest: " vẽ flow" });
    expect(leadingCommand("/clear")).toEqual({ command: "/clear", rest: "" });
    expect(leadingCommand("no command")).toBeNull();
    expect(slashBadgeRange("/sample-bd-workbook", "/sample-bd-workbook")).toEqual({ start: 0, end: 19 });
    expect(slashBadgeRange("/sample-bd-workbook ", null)).toEqual({ start: 0, end: 19 });
    expect(slashBadgeRange("/ast", null)).toBeNull();
  });
});
