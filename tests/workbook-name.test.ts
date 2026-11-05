import { describe, expect, it } from "vitest";
import { sameWorkbookName, workbookStem } from "../protocol/src/index.js";

describe("workbook names", () => {
  it("drops only the file extension", () => {
    expect(workbookStem("Book2.xlsx")).toBe("Book2");
    expect(workbookStem(" Sales 2026.XLSM ")).toBe("Sales 2026");
    expect(workbookStem("Book2")).toBe("Book2");
    expect(workbookStem("v1.2 plan.xlsx")).toBe("v1.2 plan");
  });

  it("calls two names the same when they differ only by case and extension", () => {
    expect(sameWorkbookName("Book2", "book2.xlsx")).toBe(true);
    expect(sameWorkbookName("Budget.xlsx", "Budget.xlsm")).toBe(true);
    expect(sameWorkbookName("Book2", "Book20")).toBe(false);
    expect(sameWorkbookName("Book2.xlsx", "Book1.xlsx")).toBe(false);
  });

  it("never matches a workbook with no name", () => {
    expect(sameWorkbookName("", "")).toBe(false);
    expect(sameWorkbookName(" ", "Book1")).toBe(false);
  });
});
