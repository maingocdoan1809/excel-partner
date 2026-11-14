// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { findCellRefs, isRefList, linkifyCells, mentionsFile, mentionsSheet } from "../frontend/taskpane/ui/cell-links.js";
import { cellLinkMarkdown } from "../protocol/src/index.js";
import { renderMarkdown } from "../frontend/taskpane/ui/markdown.js";

const refs = (text: string, context: "text" | "code" = "text") => findCellRefs(text, context).map((r) => `${r.sheet ?? ""}|${r.address}`);

describe("findCellRefs", () => {
  it("finds the mentions people actually write", () => {
    expect(refs("Mình đã chỉnh lại câu chữ trong F10:F16 cho tự nhiên.")).toEqual(["|F10:F16"]);
    expect(refs("Ô B5 và C7, rồi A1:B2.")).toEqual(["|B5", "|C7", "|A1:B2"]);
    expect(refs("Xem (D4), [E9] hoặc \"G1\".")).toEqual(["|D4", "|E9", "|G1"]);
    expect(refs("Totals are in $C$2:$C$9.")).toEqual(["|C2:C9"]);
    expect(refs("See AA10, XFD1048576.")).toEqual(["|AA10", "|XFD1048576"]);
  });

  it("reads the sheet when the text names one, quoted or not", () => {
    expect(refs("Data!B2")).toEqual(["Data|B2"]);
    expect(refs("Sheet1!A1:C3")).toEqual(["Sheet1|A1:C3"]);
    expect(refs("'Test Cases'!F10:F16")).toEqual(["Test Cases|F10:F16"]);
    expect(refs("'Bob''s data'!A1")).toEqual(["Bob's data|A1"]);
    expect(refs("Kết_quả!D2")).toEqual(["Kết_quả|D2"]);
    expect(findCellRefs("see 'Test Cases'!F10:F16 now")[0]).toMatchObject({ text: "'Test Cases'!F10:F16", start: 4 });
  });

  it("only links whole columns and rows when a sheet is named, so ratios and times stay text", () => {
    expect(refs("The ratio is 3:5 at 10:30")).toEqual([]);
    expect(refs("Columns D:F are hidden")).toEqual([]);
    expect(refs("Data!D:F and Data!3:5")).toEqual(["Data|D:F", "Data|3:5"]);
  });

  it("does not link words that merely look like cells", () => {
    expect(refs("Results for Q4 and H2 and FY24 improved")).toEqual([]);
    expect(refs("COVID19, B2B, a1 and sheet2 are words")).toEqual([]);
    expect(refs("version v1.2.3, 12:30pm, ID#A1")).toEqual([]);
    expect(refs("=SUM(F10:F16) and LOG10(2)")).toEqual(["|F10:F16"]);
    expect(refs("ABCD12 and ZZZZ1")).toEqual([]);
  });

  it("rejects references outside the sheet and backwards ranges", () => {
    expect(refs("XFE1 and A1048577 and A0")).toEqual([]);
    expect(refs("B2:A1 and C9:C3")).toEqual([]);
    expect(refs("Data!F:D")).toEqual([]);
  });

  it("keeps quarters and half-years linkable when they carry a sheet or sit alone in code", () => {
    expect(refs("Data!Q4")).toEqual(["Data|Q4"]);
    expect(refs("Q4", "code")).toEqual(["|Q4"]);
  });

  it("tells a list of references from a formula", () => {
    const list = (text: string) => isRefList(text, findCellRefs(text, "code"));
    expect(list("F10:F16")).toBe(true);
    expect(list(" F10:F16 ")).toBe(true);
    expect(list("C10, C14, C16")).toBe(true);
    expect(list("E10:E16; E19:E21")).toBe(true);
    expect(list("C10 và C12 and C14 or C15 hoặc C16")).toBe(true);
    expect(list("=SUM(F10:F16)")).toBe(false);
    expect(list("F10 = 5")).toBe(false);
    expect(list("Tick F10")).toBe(false);
    expect(list("no references")).toBe(false);
  });

  it("is not thrown off by punctuation next to the reference", () => {
    expect(refs("Xem F10:F16.")).toEqual(["|F10:F16"]);
    expect(refs("F10:F16, F20:F22; G1!")).toEqual(["|F10:F16", "|F20:F22", "|G1"]);
    expect(refs("trong F10:F16)")).toEqual(["|F10:F16"]);
  });
});

function render(html: string): HTMLElement {
  const root = document.createElement("div");
  root.innerHTML = html;
  linkifyCells(root);
  return root;
}
const links = (root: HTMLElement) => Array.from(root.querySelectorAll<HTMLAnchorElement>("a.cell-link")).map((a) => [a.textContent, a.dataset.sheet ?? "", a.dataset.address]);

describe("linkifyCells", () => {
  it("turns the user's example sentence into one link and keeps the rest of the text", () => {
    const root = render("<p>Mình đã chỉnh lại câu chữ trong F10:F16 cho tự nhiên và nhất quán hơn.</p>");
    expect(links(root)).toEqual([["F10:F16", "", "F10:F16"]]);
    expect(root.textContent).toBe("Mình đã chỉnh lại câu chữ trong F10:F16 cho tự nhiên và nhất quán hơn.");
    expect(root.querySelector("a")!.getAttribute("href")).toBe("#");
  });

  it("sends a reference to the sheet named before it in the same text", () => {
    const root = render("<p>In Data!B2 I wrote 5, then C3 and D4. Later 'Test Cases'!A1, then E5.</p>");
    expect(links(root)).toEqual([
      ["Data!B2", "Data", "B2"], ["C3", "Data", "C3"], ["D4", "Data", "D4"], ["'Test Cases'!A1", "Test Cases", "A1"], ["E5", "Test Cases", "E5"],
    ]);
  });

  it("leaves a reference with no sheet in the text for the open sheet", () => {
    expect(links(render("<p>Check B2</p>"))).toEqual([["B2", "", "B2"]]);
  });

  it("links inline code only when it is nothing but references", () => {
    const root = render("<p>Use <code>F10:F16</code> not <code>=SUM(F10:F16)</code> or <code>Q4</code>.</p>");
    expect(links(root)).toEqual([["F10:F16", "", "F10:F16"], ["Q4", "", "Q4"]]);
    expect(root.querySelectorAll("code")[0]!.textContent).toBe("=SUM(F10:F16)");
  });

  it("links every cell in a code span that lists them, and drops the code box", () => {
    const root = render("<p>Các ô đã tick: <code>C10, C14, C16</code> và <code>E11, E23</code>.</p>");
    expect(links(root)).toEqual([["C10", "", "C10"], ["C14", "", "C14"], ["C16", "", "C16"], ["E11", "", "E11"], ["E23", "", "E23"]]);
    expect(root.querySelector("code")).toBeNull();
    expect(root.querySelector("p")!.textContent).toBe("Các ô đã tick: C10, C14, C16 và E11, E23.");
  });

  it("replaces the code box with the badge, so there is no second bubble behind it", () => {
    const root = render("<p>Sửa <code>F10:F16</code> và <code>Data!B2</code> rồi <code>=A1+B1</code>.</p>");
    expect(root.querySelector("code a")).toBeNull();
    expect(root.querySelectorAll("code")).toHaveLength(1);
    expect(root.querySelectorAll("p > a.cell-link")).toHaveLength(2);
    expect(root.querySelector("p")!.textContent).toBe("Sửa F10:F16 và Data!B2 rồi =A1+B1.");
  });

  it("leaves code blocks and existing links alone", () => {
    const root = render('<pre><code>F10:F16 = 5</code></pre><p><a href="https://example.com/A1">A1 docs</a> and B2</p>');
    expect(links(root)).toEqual([["B2", "", "B2"]]);
    expect(root.querySelector("pre a")).toBeNull();
    expect(root.querySelector('a[href^="https"]')!.textContent).toBe("A1 docs");
  });

  it("works inside lists, tables and emphasis", () => {
    const root = render("<ul><li><strong>F10</strong>: sửa</li></ul><table><tr><td>G1:G3</td></tr></table>");
    expect(links(root)).toEqual([["F10", "", "F10"], ["G1:G3", "", "G1:G3"]]);
  });

  it("cannot be used to inject markup through a sheet name", () => {
    const root = document.createElement("div");
    root.textContent = `'x"><img src=x onerror=alert(1)>'!A1 and Data!B2`;
    linkifyCells(root);
    expect(root.querySelector("img")).toBeNull();
    const link = root.querySelector<HTMLAnchorElement>("a.cell-link")!;
    expect(root.querySelectorAll("a.cell-link").length).toBeGreaterThan(0);
    expect(link.querySelector("img")).toBeNull();
  });

  it("keeps a sheet name with quotes and markup characters as plain data", () => {
    const root = document.createElement("div");
    root.textContent = "'a''<b>'!C3";
    linkifyCells(root);
    const link = root.querySelector<HTMLAnchorElement>("a.cell-link")!;
    expect(link.dataset.sheet).toBe("a'<b>");
    expect(root.querySelector("b")).toBeNull();
  });

  it("does nothing when there is nothing to link", () => {
    const root = render("<p>Không có ô nào ở đây, chỉ có Q4 và 12:30.</p>");
    expect(root.querySelector("a")).toBeNull();
    expect(root.innerHTML).toBe("<p>Không có ô nào ở đây, chỉ có Q4 và 12:30.</p>");
  });
});

describe("renderMarkdown with cellLinks", () => {
  it("is off unless asked for", () => {
    expect(renderMarkdown("Xem F10:F16")).not.toContain("cell-link");
  });

  it("links references in a chat reply and keeps real web links safe", () => {
    const html = renderMarkdown("Mình sửa **F10:F16** và `Data!B2`, xem [docs](https://example.com) và [x](javascript:alert(1)).", { cellLinks: true });
    const root = document.createElement("div");
    root.innerHTML = html;
    expect(links(root)).toEqual([["F10:F16", "", "F10:F16"], ["Data!B2", "Data", "B2"]]);
    expect(root.querySelector('a[href^="https://example.com"]')!.getAttribute("rel")).toBe("noopener noreferrer");
    expect(html).not.toContain("javascript:");
  });
});

describe("links the assistant spells out", () => {
  const draw = (markdown: string, cellLinks: boolean | { plain?: boolean } = true) => {
    const root = document.createElement("div");
    root.innerHTML = renderMarkdown(markdown, { cellLinks });
    return root;
  };
  const where = (root: HTMLElement) => Array.from(root.querySelectorAll<HTMLAnchorElement>("a.cell-link"))
    .map((a) => `${a.textContent} -> ${a.dataset.file ?? "-"} | ${a.dataset.sheet ?? "-"} | ${a.dataset.address}`);

  it("goes exactly where the link says: the sentence that was wrong before", () => {
    const root = draw("Mình đã nhập dữ liệu vào **Book2 → Sheet1**, vùng [A1:E8](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8) (bảng bán hàng mẫu):");
    expect(where(root)).toEqual(["A1:E8 -> Book2.xlsx | Sheet1 | A1:E8"]);
    expect(root.querySelector<HTMLAnchorElement>("a.cell-link")!.title).toBe("Select Book2.xlsx → Sheet1!A1:E8 in Excel");
    expect(root.textContent?.trim()).toBe("Mình đã nhập dữ liệu vào Book2 → Sheet1, vùng A1:E8 (bảng bán hàng mẫu):");
  });

  it("keeps the file, sheet and range intact through the markdown and back, whatever the names hold", () => {
    for (const target of [
      { file: "Sales (2026) báo cáo.xlsx", sheet: "Bảng tổng hợp [Q1]", address: "B2:C3" },
      { file: "Book 2.xlsx", sheet: "Sheet's & more", address: "D:D" },
      { sheet: "Dữ liệu", address: "3:5" },
      { address: "AB12" },
    ]) {
      const link = draw(cellLinkMarkdown(target.address, target)).querySelector<HTMLAnchorElement>("a.cell-link")!;
      expect({ file: link.dataset.file, sheet: link.dataset.sheet, address: link.dataset.address }).toEqual({ file: target.file, sheet: target.sheet, address: target.address });
    }
  });

  it("reads a name the assistant wrote with accents and no percent-encoding", () => {
    expect(where(draw("[x](cell:?sheet=Bảng%20tổng&range=a1)"))).toEqual(["x -> - | Bảng tổng | A1"]);
  });

  it("leaves a link without a file in this workbook", () => {
    expect(where(draw("[B2](cell:?sheet=Data&range=B2) và [C3](cell:?range=C3)"))).toEqual(["B2 -> - | Data | B2", "C3 -> - | - | C3"]);
  });

  it("shows only the words of a cell link that names no real range, and of one that is not a cell link at all", () => {
    for (const href of ["cell:?range=ZZZZ1", "cell:?sheet=Data", "cell:?range=A1;DROP", "cell:?range=B5:A1", "cell:?range=A1:B2:C3", "cell:", "cell:?range=A0"]) {
      const root = draw(`See [here](${href}) now.`);
      expect(root.querySelector("a"), href).toBeNull();
      expect(root.textContent?.trim(), href).toBe("See here now.");
    }
    const unsafe = draw("[x](javascript:alert(1)) and [y](data:text/html,<b>) and [z](https://example.com)");
    expect(unsafe.querySelectorAll("a")).toHaveLength(1);
    expect(unsafe.querySelector("a")!.getAttribute("href")).toBe("https://example.com");
  });

  it("is text when links are off", () => {
    const root = draw("Xem [A1:E8](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8)", false);
    expect(root.querySelector("a")).toBeNull();
    expect(root.textContent?.trim()).toBe("Xem A1:E8");
  });

  it("does not link a second time inside a badge it made", () => {
    const root = draw("[see B2:C3](cell:?range=B2:C3)");
    expect(where(root)).toEqual(["see B2:C3 -> - | - | B2:C3"]);
    expect(root.querySelectorAll("a")).toHaveLength(1);
  });

  it("guesses nothing from the words around bare references when it is told the reply may be about another file", () => {
    const text = "Mình đã nhập dữ liệu vào **Book2 → Sheet1**, vùng A1:E8, và Sheet1!B2 nữa; xem [C3](cell:?file=Book2.xlsx&sheet=Sheet1&range=C3).";
    expect(where(draw(text, { plain: false }))).toEqual(["C3 -> Book2.xlsx | Sheet1 | C3"]);
    // Otherwise bare references link as before: a file is never part of one.
    expect(where(draw(text, { plain: true }))).toEqual(["A1:E8 -> - | - | A1:E8", "Sheet1!B2 -> - | Sheet1 | B2", "C3 -> Book2.xlsx | Sheet1 | C3"]);
  });
});

describe("mentionsFile", () => {
  const files = ["Book2.xlsx", "Sales (2026).xlsx"];

  it("finds a file by its name, any case, with or without the extension", () => {
    expect(mentionsFile("Đã nhập vào Book2 → Sheet1", files)).toBe(true);
    expect(mentionsFile("saved in BOOK2.XLSX", files)).toBe(true);
    expect(mentionsFile("see Sales (2026) for it", files)).toBe(true);
  });

  it("does not take part of another word for a file", () => {
    expect(mentionsFile("Book20 and Notebook2", files)).toBe(false);
    expect(mentionsFile("nothing here", files)).toBe(false);
    expect(mentionsFile("Book2", [])).toBe(false);
  });

  it("ignores names too short to tell from ordinary words", () => {
    expect(mentionsFile("a b c", ["A.xlsx", ""])).toBe(false);
  });
});

describe("mentionsSheet", () => {
  const sheets = ["Sheet1", "Sheet2", "Tổng hợp", "Data (2026)", "A"];

  it("finds a sheet by its name, any case, wherever the sentence puts it", () => {
    expect(mentionsSheet("mình chưa đụng gì vào workbook Excel — sheet2/C5 vẫn nguyên", sheets)).toBe(true);
    expect(mentionsSheet("Trong Tổng hợp, ô B2 sai", sheets)).toBe(true);
    expect(mentionsSheet("xem DATA (2026): C3", sheets)).toBe(true);
  });

  it("does not take part of another word, or a one-letter name, for a sheet", () => {
    expect(mentionsSheet("Sheet20 and MySheet2", sheets)).toBe(false);
    expect(mentionsSheet("a total in a cell, C5", sheets)).toBe(false);
    expect(mentionsSheet("ô C5 vẫn nguyên", sheets)).toBe(false);
    expect(mentionsSheet("anything", [])).toBe(false);
  });

  it("is what keeps a bare C5 from being linked to the open sheet when the sentence says sheet2", () => {
    const text = "mình chưa đụng gì vào workbook Excel — sheet2/C5 vẫn nguyên.";
    const draw = (plain: boolean) => {
      const root = document.createElement("div");
      root.innerHTML = renderMarkdown(text, { cellLinks: { plain } });
      return root.querySelectorAll("a.cell-link").length;
    };
    expect(draw(true)).toBe(1); // what the page used to do: C5 on whatever sheet is open
    expect(draw(!mentionsSheet(text, sheets))).toBe(0);
  });
});
