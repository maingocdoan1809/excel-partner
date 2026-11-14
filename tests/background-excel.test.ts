import { afterEach, describe, expect, it, vi } from "vitest";
import { executeOfficeJs } from "../frontend/taskpane/executor/office-js.js";
import { installOfficeMock, MockWorkbook } from "./support/office-mock.js";

afterEach(() => vi.unstubAllGlobals());

function fixture() {
  const range = { values: [[0]], select: vi.fn(), copyFrom: vi.fn(), load() { return this; } };
  const sheet = { activate: vi.fn(), getRange: vi.fn(() => range) };
  const sync = vi.fn(async () => undefined);
  const context = { sync, workbook: { worksheets: { getItem: vi.fn(() => sheet), items: [sheet] } },
    application: { suspendScreenUpdatingUntilNextSync: vi.fn() } };
  const run = vi.fn(async (callback: (ctx: typeof context) => Promise<unknown>) => callback(context));
  vi.stubGlobal("Excel", { run });
  vi.stubGlobal("Office", { context: {} });
  return { range, sheet, context, run };
}

describe("generated Excel updates leave the user's viewport alone", () => {
  it("writes to a named sheet while another sheet stays active, with loaded Office proxies", async () => {
    const book = new MockWorkbook();
    book.addSheet("Reading");
    book.addSheet("Report");
    installOfficeMock(book);
    const result = await executeOfficeJs({ code: `
      const sheet = context.workbook.worksheets.getItem('Report');
      sheet.activate();
      sheet.getRange('A1:B1').values = [[12, 24]];
      const range = sheet.getRange('A1:B1');
      range.load('values');
      await context.sync();
      return range.values;
    ` });
    expect(result).toMatchObject({ result: [[12, 24]], navigationSuppressed: ["activate"] });
    expect(book.active).toBe(0);
    expect(book.get("Report", "B1")).toBe(24);
    expect(book.get("Reading", "B1")).toBeNull();
  });

  it("refuses activate-then-write-to-active-sheet instead of writing into the user's sheet", async () => {
    const book = new MockWorkbook();
    book.addSheet("Reading"); book.addSheet("Report"); installOfficeMock(book);
    await expect(executeOfficeJs({ code: `
      context.workbook.worksheets.getItem('Report').activate();
      context.workbook.worksheets.getActiveWorksheet().getRange('A1').values = [[99]];
    ` })).rejects.toThrow("Use explicit worksheet and range objects");
    expect(book.active).toBe(0);
    expect(book.get("Reading", "A1")).toBeNull();
    expect(book.get("Report", "A1")).toBeNull();
  });

  it("writes and returns results without activating, selecting or repeatedly freezing the screen", async () => {
    const f = fixture();
    const result = await executeOfficeJs({ code: `
      const sheet = context.workbook.worksheets.getItem('Report');
      sheet.activate();
      const range = sheet.getRange('A1');
      range.load('values');
      range.select();
      for (let i = 1; i <= 3; i++) {
        context.application.suspendScreenUpdatingUntilNextSync();
        range.values = [[i]];
        await context.sync();
      }
      return range.values;
    ` });
    expect(f.range.values).toEqual([[3]]);
    expect(result).toMatchObject({ result: [[3]], navigationSuppressed: ["activate", "select", "suspendScreenUpdatingUntilNextSync"] });
    expect(f.sheet.activate).not.toHaveBeenCalled();
    expect(f.range.select).not.toHaveBeenCalled();
    expect(f.context.application.suspendScreenUpdatingUntilNextSync).not.toHaveBeenCalled();
  });

  it("unwraps Office object arguments and protects collection items and nested batches", async () => {
    const f = fixture();
    await executeOfficeJs({ code: `
      context.workbook.worksheets.items.forEach(sheet => sheet.activate());
      const range = context.workbook.worksheets.getItem('Report').getRange('A1');
      range.copyFrom(range);
      await Excel.run(async other => { other.workbook.worksheets.getItem('Report').activate(); });
    ` });
    expect(f.range.copyFrom).toHaveBeenCalledWith(f.range);
    expect(f.sheet.activate).not.toHaveBeenCalled();
    expect(f.run).toHaveBeenCalledTimes(2);
  });

  it("does not disable real user navigation while the generated batch awaits sync", async () => {
    const f = fixture();
    f.context.sync.mockImplementationOnce(async () => { f.sheet.activate(); f.range.select(); });
    await executeOfficeJs({ code: `
      const sheet = context.workbook.worksheets.getItem('Report');
      sheet.activate();
      await context.sync();
      sheet.getRange('A1').values = [[7]];
      sheet.activate();
    ` });
    expect(f.sheet.activate).toHaveBeenCalledTimes(1);
    expect(f.range.select).toHaveBeenCalledTimes(1);
    expect(f.range.values).toEqual([[7]]);
    // Nothing was patched globally, even after an exception.
    await expect(executeOfficeJs({ code: "throw new Error('failure')" })).rejects.toThrow("failure");
    f.sheet.activate();
    expect(f.sheet.activate).toHaveBeenCalledTimes(2);
  });
});
