import { beforeEach, describe, expect, it } from "vitest";
import { executeTool } from "../addin/taskpane/executor/index.js";
import { describeQaApproval } from "../addin/taskpane/permissions.js";
import { describeChange } from "../addin/taskpane/ui/changes.js";
import { WRITE_TOOLS } from "../addin/taskpane/executor/index.js";
import { MockWorkbook, installOfficeMock } from "./support/office-mock.js";

const COLUMNS = { tcId: "TC_ID", status: "Automation_Status", scriptPath: "Script_Path", lastRun: "Last_Run", lastResult: "Last_Result", duration: "Duration_s", evidence: "Run_Evidence" };
let book: MockWorkbook;

beforeEach(() => {
  book = new MockWorkbook();
  book.addSheet("TestCases");
  book.fill("TestCases", [
    ["TC_ID", "Title", "Automation_Status", "Notes"],
    ["TC_001", "Login", "Manual", "keep me"],
    ["TC_002", "Logout", null, null],
    ["tc_003", "Export", "Manual", "also keep"],
  ]);
  installOfficeMock(book);
});

const apply = (rows: Array<{ tcId: string; values: Record<string, string | number | null> }>, extra: Record<string, unknown> = {}) =>
  executeTool("__qa_apply_results", { sheet: "TestCases", columns: COLUMNS, rows, ...extra });

describe("writing QA results into the sheet", () => {
  it("matches rows by TC_ID (case-insensitive), fills existing columns and appends missing ones", async () => {
    const result = await apply([
      { tcId: "TC_001", values: { status: "Approved", scriptPath: "automation/tests/TC_001_login.spec.js", lastResult: "Passed", duration: 1.2 } },
      { tcId: "TC_003", values: { status: "Draft", lastResult: "Failed", duration: 3 } },
    ]);
    expect(result).toMatchObject({ ok: true, data: { sheet: "TestCases", rowsUpdated: 2, notInSheet: [], headerRow: 1 } });
    expect((result as any).data.addedColumns).toEqual(["Script_Path", "Last_Result", "Duration_s"]);
    // Existing "Automation_Status" column is reused; the rest are appended after the last header (D -> E, F, G).
    expect(book.get("TestCases", "C2")).toBe("Approved");
    expect(book.get("TestCases", "C4")).toBe("Draft");
    expect(book.get("TestCases", "E1")).toBe("Script_Path");
    expect(book.get("TestCases", "F1")).toBe("Last_Result");
    expect(book.get("TestCases", "G1")).toBe("Duration_s");
    expect(book.get("TestCases", "E2")).toBe("automation/tests/TC_001_login.spec.js");
    expect(book.get("TestCases", "F4")).toBe("Failed");
    expect(Number(book.get("TestCases", "G2"))).toBe(1.2);
  });

  it("leaves other cells, other rows and rows without a result untouched", async () => {
    await apply([{ tcId: "TC_001", values: { status: "Approved" } }]);
    expect(book.get("TestCases", "D2")).toBe("keep me");
    expect(book.get("TestCases", "D4")).toBe("also keep");
    expect(book.get("TestCases", "C3")).toBeNull();
    expect(book.get("TestCases", "C4")).toBe("Manual");
    expect(book.get("TestCases", "B3")).toBe("Logout");
  });

  it("does not create columns it has no values for", async () => {
    const result = await apply([{ tcId: "TC_001", values: { status: "Approved" } }]);
    expect((result as any).data.addedColumns).toEqual([]);
    expect(book.get("TestCases", "E1")).toBeNull();
  });

  it("reports scripts that have no row in the sheet", async () => {
    const result = await apply([{ tcId: "TC_404", values: { status: "Draft" } }, { tcId: "TC_002", values: { status: "Draft" } }]);
    expect((result as any).data).toMatchObject({ rowsUpdated: 1, notInSheet: ["TC_404"] });
    expect((result as any).data.note).toContain("TC_404");
  });

  it("keeps formulas in a column it fills for rows it does not touch", async () => {
    book.sheet("TestCases").setCell(2, 2, { v: 5, f: "=1+4" });
    await apply([{ tcId: "TC_001", values: { status: "Approved" } }]);
    expect(book.get("TestCases", "C3")).toBe("=1+4");
  });

  it("never turns a result into a formula", async () => {
    await apply([{ tcId: "TC_001", values: { status: "=HYPERLINK(\"http://x\")", scriptPath: "+cmd", lastResult: "-1+1", evidence: "@SUM(A1)" } }]);
    for (const address of ["C2", "E2", "F2", "G2"]) expect(String(book.get("TestCases", address)).startsWith("'")).toBe(true);
  });

  it("finds the header row below a title block", async () => {
    book = new MockWorkbook();
    book.addSheet("TestCases");
    book.fill("TestCases", [["Project X test plan", null], [null, null], ["TC_ID", "Title"], ["TC_001", "Login"]]);
    installOfficeMock(book);
    const result = await apply([{ tcId: "TC_001", values: { status: "Draft" } }]);
    expect(result).toMatchObject({ ok: true, data: { headerRow: 3, rowsUpdated: 1 } });
    expect(book.get("TestCases", "C3")).toBe("Automation_Status");
    expect(book.get("TestCases", "C4")).toBe("Draft");
  });

  it("explains what is wrong when the sheet or the TC_ID column is missing", async () => {
    expect(await apply([], { sheet: "Nope" })).toMatchObject({ ok: false, error: { code: "sheet_not_found" } });
    book.addSheet("Empty");
    expect(await apply([], { sheet: "Empty" })).toMatchObject({ ok: false, error: { code: "range_not_found" } });
    book.addSheet("NoId");
    book.fill("NoId", [["A", "B"], [1, 2]]);
    expect(await apply([], { sheet: "NoId" })).toMatchObject({ ok: false, error: { code: "range_not_found", message: expect.stringContaining("TC_ID") } });
  });

  it("supports a customised header text", async () => {
    book.fill("TestCases", [["Case no"]], "A1");
    const result = await apply([{ tcId: "TC_001", values: { status: "Draft" } }], { columns: { ...COLUMNS, tcId: "Case no" } });
    expect(result).toMatchObject({ ok: true, data: { rowsUpdated: 1 } });
  });

  it("is listed as a workbook change with a readable summary", async () => {
    expect(WRITE_TOOLS.has("__qa_apply_results")).toBe(true);
    const result = await apply([{ tcId: "TC_001", values: { status: "Approved", lastResult: "Passed" } }]);
    const change = describeChange("__qa_apply_results", {}, result, WRITE_TOOLS);
    expect(change!.summary).toBe("QA results written to TestCases: 1 test case(s), added column(s) Last_Result");
  });

  it("rejects malformed payloads before touching the workbook", async () => {
    expect(await executeTool("__qa_apply_results", { sheet: "TestCases", columns: COLUMNS, rows: [{ tcId: "TC_001", values: { x: { evil: true } } }] })).toMatchObject({ ok: false });
    expect(book.get("TestCases", "C2")).toBe("Manual");
  });
});

describe("QA approval dialogs", () => {
  it("shows the code that will be saved and says it stays a draft", () => {
    const text = describeQaApproval({ action: "save", kind: "test", id: "TC_001", file: "automation/tests/TC_001_x.spec.js", code: "// TC: TC_001\n" });
    expect(text.target).toContain("// TC: TC_001");
    expect(text.detail).toContain("qa approve TC_001");
    expect(text.confirm).toBe("Save draft");
  });

  it("warns plainly about unapproved scripts and lists them", () => {
    const text = describeQaApproval({ action: "run_draft", forceAsk: true, environment: "staging", baseUrl: "https://s.test", scripts: [{ tcId: "TC_002", title: "Refund", role: "admin", status: "draft" }] });
    expect(text.title).toContain("UNAPPROVED");
    expect(text.detail).toContain("nobody has approved");
    expect(text.detail).toContain("only that site");
    expect(text.target).toContain("TC_002  Refund  [admin, draft]");
  });

  it("states the sites a signed-in browser may reach", () => {
    const text = describeQaApproval({ action: "open_session", profile: "admin", environment: "staging", origins: ["https://s.test", "https://sso.test"] });
    expect(text.detail).toContain("https://s.test, https://sso.test");
    expect(text.detail).toContain("password is never shared");
  });
});
