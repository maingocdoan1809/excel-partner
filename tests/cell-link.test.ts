import { describe, expect, it } from "vitest";
import { cellHref, cellLinkMarkdown, isA1Range, parseCellHref } from "../protocol/src/index.js";

describe("cell: links", () => {
  it("writes the address readably and encodes only what must be", () => {
    expect(cellHref({ file: "Book2.xlsx", sheet: "Sheet1", address: "A1:E8" })).toBe("cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8");
    expect(cellHref({ file: "Sales (2026).xlsx", sheet: "My Sheet", address: "$B$2" })).toBe("cell:?file=Sales%20(2026).xlsx&sheet=My%20Sheet&range=B2");
    expect(cellHref({ address: "D:D" })).toBe("cell:?range=D:D");
  });

  it("reads back exactly what was written, for any name", () => {
    for (const target of [
      { file: "Book2.xlsx", sheet: "Sheet1", address: "A1:E8" },
      { file: "Sales (2026) báo cáo.xlsx", sheet: "Bảng [Q1] & 'x'", address: "B2" },
      { sheet: "Dữ liệu", address: "3:5" },
      { address: "XFD1048576" },
    ]) expect(parseCellHref(cellHref(target))).toEqual(target);
  });

  it("escapes the label so a bracket cannot end the link early", () => {
    expect(cellLinkMarkdown("a]b[c", { address: "A1" })).toBe("[a\\]b\\[c](cell:?range=A1)");
  });

  it("refuses what is not a cell link naming a real range", () => {
    for (const href of ["", "https://example.com", "cell:", "cell:?sheet=Data", "cell:?range=", "cell:?range=A0", "cell:?range=XFE1",
      "cell:?range=A1:B", "cell:?range=B2:A1", "cell:?range=D:A", "cell:?range=5:3", "cell:?range=A1;B2", "cell:?range=Sheet1!A1",
      `cell:?file=${"x".repeat(300)}&range=A1`]) {
      expect(parseCellHref(href), href).toBeNull();
    }
  });

  it("takes lower case, dollar signs and surrounding space", () => {
    expect(parseCellHref(" cell:?range=%24a%241:%24b%242 ")).toEqual({ address: "A1:B2" });
    expect(parseCellHref("CELL:?range=b2")).toEqual({ address: "B2" });
  });

  it("knows a range when it sees one", () => {
    for (const ok of ["A1", "AB12", "A1:E8", "D:D", "A:C", "3:5", "XFD1048576"]) expect(isA1Range(ok), ok).toBe(true);
    for (const no of ["", "A", "1", "A1:", "A1:B2:C3", "a1", "AAAA1", "A1048577", "C:A", "A1 "]) expect(isA1Range(no), no).toBe(false);
  });
});
