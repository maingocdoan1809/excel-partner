import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";
import { exposedToolNames, parseToolInput } from "../protocol/src/index.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { TestScripts, validateScript } from "../bridge/src/test-scripts.js";

const origin = "https://staging.example.com";
const goodScript = `import { test, expect } from '@playwright/test';
test.use({ baseURL: process.env.BASE_URL, storageState: process.env.STORAGE_STATE });
test('TC-001 create ticket', async ({ page }) => {
  await page.goto('/tickets/new');
  await page.getByLabel('Title').fill('Printer broken');
  await page.getByRole('button', { name: 'Save' }).click();
  await expect(page.getByText('Ticket created')).toBeVisible();
});`;
const save = (overrides: Record<string, unknown> = {}) => ({
  origin, caseId: "TC-001", title: "Create a ticket", sheet: "Test Cases", row: 4, outcome: "pass",
  evidence: "Saved the ticket and the green banner 'Ticket created' appeared.", script: goodScript, ...overrides,
});

function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ep-scripts-"));
  const scripts = new TestScripts(dir);
  const sessions = new SessionManager({ ttlMs: 60_000 });
  const session = sessions.create("scripts", null);
  session.capabilities.playwright = true;
  Object.defineProperty(session, "isConnected", { get: () => true, configurable: true });
  const gateway = new ToolGateway(sessions, { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 }, 1000,
    undefined, false, null, null, null, undefined, null, null, scripts);
  /** A fake task pane: a one-row-per-case sheet that keeps what record_test_results writes. */
  const cells = new Map<string, string>();
  const ask = vi.spyOn(session, "requestTool").mockImplementation(async (tool: string, args: unknown) => {
    if (tool === "record_test_results") {
      const input = args as { results: Array<{ row: number; values: Record<string, string> }> };
      const columns: Record<string, string> = {};
      for (const item of input.results) for (const [header, value] of Object.entries(item.values)) {
        const letter = header === "Tiến độ" ? "F" : "G";
        columns[header] = letter;
        cells.set(`${letter}${item.row}`, String(value));
      }
      return { ok: true, data: { columns } };
    }
    if (tool === "read_range") {
      const row = Number((args as { address: string }).address.match(/(\d+)$/)![1]);
      return { ok: true, data: { values: [["", "", "", "", "", cells.get(`F${row}`) ?? "", cells.get(`G${row}`) ?? ""]] } };
    }
    return { ok: true, data: { approved: true } };
  });
  const tool = (name: string, args: unknown) => gateway.invoke({ sessionId: session.id, tool: name, args });
  const record = (row: number, status: string) => tool("record_test_results", { sheet: "Test Cases", headerRow: 3, results: [{ row, values: { "Tiến độ": status, "Kết quả thực tế": "ok" } }] });
  return { dir, scripts, session, ask, cells, tool, record };
}

describe("regression script rules", () => {
  it("accepts a Playwright Test file that checks something and rejects everything else", () => {
    expect(() => validateScript(origin, goodScript)).not.toThrow();
    const bad: Array<[string, string]> = [
      ["no import", goodScript.replace(/^import .*\n/, "")],
      ["other import", `import fs from 'node:fs';\n${goodScript}`],
      ["no check", goodScript.replace(/await expect.*\n/, "")],
      ["require", goodScript.replace("await page.goto('/tickets/new');", "const x = require('fs');")],
      ["fetch", goodScript.replace("await page.goto('/tickets/new');", "await fetch('/x');")],
      ["other process value", goodScript.replace("process.env.STORAGE_STATE", "process.env.HOME")],
      ["eval", goodScript.replace("await page.goto('/tickets/new');", "eval('1');")],
      ["another site", goodScript.replace("'/tickets/new'", "'https://evil.example.org/x'")],
      ["a password", goodScript.replace("'Printer broken'", "'x'); const password = 'hunter2hunter2'; (void 0")],
    ];
    for (const [name, script] of bad) expect(() => validateScript(origin, script), name).toThrow();
    expect(() => validateScript(origin, goodScript.replace("'/tickets/new'", "'https://staging.example.com/tickets/new'"))).not.toThrow();
  });

  it("is only offered together with the browser tools", () => {
    expect(exposedToolNames(false, { playwright: true })).toEqual(expect.arrayContaining(["save_test_script", "finish_test_run"]));
    expect(exposedToolNames(false, {})).not.toContain("save_test_script");
    expect(() => parseToolInput("save_test_script", save({ outcome: "blocked" }))).toThrow();
    expect(() => parseToolInput("save_test_script", save({ evidence: "short" }))).toThrow();
    expect(() => parseToolInput("save_test_script", save({ caseId: "../x" }))).toThrow();
  });
});

describe("saving scripts", () => {
  it("writes a readable file with a header, lists it, replaces it and deletes it", async () => {
    const { dir, scripts, tool } = setup();
    expect(await tool("save_test_script", save())).toMatchObject({ ok: true, data: { saved: { caseId: "TC-001", outcome: "pass" } } });
    const file = path.join(dir, "staging.example.com", "TC-001.spec.js");
    const text = fs.readFileSync(file, "utf8");
    expect(text).toContain("// outcome: pass");
    expect(text).toContain("// sheet: Test Cases (row 4)");
    expect(text).toContain("import { test, expect } from '@playwright/test';");
    expect(scripts.list(origin)).toEqual([expect.objectContaining({ caseId: "TC-001", title: "Create a ticket", outcome: "pass" })]);
    await tool("save_test_script", save({ outcome: "fail_app_defect", title: "Create a ticket (bug)" }));
    expect(scripts.list(origin)).toEqual([expect.objectContaining({ outcome: "fail_app_defect", title: "Create a ticket (bug)" })]);
    expect(fs.readFileSync(file, "utf8")).toContain("supposed to fail until the application defect is fixed");
    expect(scripts.delete(origin, "TC-001")).toBe(true);
    expect(scripts.list(origin)).toEqual([]);
  });

  it("refuses an unsafe script and keeps nothing on disk", async () => {
    const { dir, tool } = setup();
    expect(await tool("save_test_script", save({ script: goodScript.replace("await page.goto('/tickets/new');", "await fetch('https://x.test');") })))
      .toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(fs.existsSync(path.join(dir, "staging.example.com"))).toBe(false);
  });

  it("keeps a multi-line title and evidence on one comment line so they cannot add code", async () => {
    const { dir, tool } = setup();
    await tool("save_test_script", save({ title: "Create\nawait page.evaluate('x')", evidence: "Saw the banner.\nawait evil();\nand more text here" }));
    const lines = fs.readFileSync(path.join(dir, "staging.example.com", "TC-001.spec.js"), "utf8").split("\n").filter(line => line.includes("evil") || line.includes("page.evaluate('x')"));
    expect(lines.every(line => line.startsWith("//"))).toBe(true);
  });
});

describe("finishing a run", () => {
  const finish = (tool: ReturnType<typeof setup>["tool"], cases: Array<Record<string, unknown>>) => tool("finish_test_run", { sheet: "Test Cases", cases });

  it("says done only when results are in the sheet and kept cases have their script", async () => {
    const { tool, record } = setup();
    await record(4, "Pass");
    await tool("save_test_script", save());
    await record(5, "Blocked");
    expect(await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "pass" }, { row: 5, caseId: "TC-002", outcome: "blocked" }]))
      .toMatchObject({ ok: true, data: { done: true, cases: 2, pass: 1, blocked: 1, scriptsSaved: 1 } });
  });

  it("lists what is missing: no result, no script, a script that should not exist", async () => {
    const { tool, record } = setup();
    const missing = await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "pass" }]);
    expect(missing).toMatchObject({ ok: false, error: { code: "invalid_arguments", detail: { problems: [expect.stringContaining("no result has been recorded"), expect.stringContaining("needs a regression script")] } } });
    await record(4, "Pass");
    expect(await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "pass" }])).toMatchObject({ ok: false, error: { detail: { problems: [expect.stringContaining("needs a regression script")] } } });
    await tool("save_test_script", save());
    expect(await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "needs_review" }])).toMatchObject({ ok: false, error: { detail: { problems: [expect.stringContaining("a script was saved although")] } } });
    expect(await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "fail_app_defect" }])).toMatchObject({ ok: false, error: { detail: { problems: [expect.stringContaining("not TC-001 / fail_app_defect")] } } });
  });

  it("reads the row back and notices a cell that no longer holds the recorded status", async () => {
    const { tool, record, cells } = setup();
    await record(4, "Pass");
    await tool("save_test_script", save());
    cells.set("F4", "Đang test");
    const result = await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "pass" }]);
    expect(result).toMatchObject({ ok: false, error: { detail: { problems: [expect.stringContaining("should hold 'Pass' but holds 'Đang test'")] } } });
    cells.set("F4", "Pass");
    expect(await finish(tool, [{ row: 4, caseId: "TC-001", outcome: "pass" }])).toMatchObject({ ok: true });
  });

  it("is not satisfied by a result the sheet refused", async () => {
    const { tool, session, ask } = setup();
    const original = ask.getMockImplementation()!;
    ask.mockImplementation(async (name, args, timeout) => name === "record_test_results" ? { ok: false, error: { code: "would_overwrite", message: "no" } } : original(name, args, timeout));
    await tool("record_test_results", { sheet: "Test Cases", headerRow: 3, results: [{ row: 4, values: { "Tiến độ": "Pass" } }] });
    expect(session.testRun.recorded.size).toBe(0);
  });
});
