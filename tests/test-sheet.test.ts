import { beforeEach, describe, expect, it, vi } from "vitest";
import { executeTool, WRITE_TOOLS } from "../addin/taskpane/executor/index.js";
import { describeTestBrowserApproval } from "../addin/taskpane/permissions.js";
import { describeChange } from "../addin/taskpane/ui/changes.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { TestSessions } from "../bridge/src/test-sessions.js";
import { SiteKnowledge, siteKey } from "../bridge/src/site-knowledge.js";
import { describeSiteKnowledgeApproval } from "../addin/taskpane/permissions.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { exposedToolNames, normalizeOrigin, parseToolInput } from "../protocol/src/index.js";
import { MockWorkbook, installOfficeMock } from "./support/office-mock.js";

let book: MockWorkbook;

/** A company-style layout: a title block above the header, headers that are not at row 1, a merged-looking ID column. */
beforeEach(() => {
  book = new MockWorkbook();
  book.addSheet("Test Cases");
  book.fill("Test Cases", [
    ["Project X - Test cases", null, null, null],
    [null, null, null, null],
    ["No.", "Chức năng", "Bước thực hiện", "Kết quả mong đợi"],
    ["UC01-001", "Đăng nhập", "Nhập user/pass hợp lệ", "Vào trang chủ"],
    ["UC01-002", "Đăng nhập", "Nhập sai mật khẩu", "Báo lỗi"],
    ["UC01-003", "Đăng xuất", "Bấm Đăng xuất", "Về trang đăng nhập"],
  ]);
  installOfficeMock(book);
});

const record = (args: Record<string, unknown>) =>
  executeTool("record_test_results", parseToolInput("record_test_results", { sheet: "Test Cases", headerRow: 3, ...args }));

describe("record_test_results", () => {
  it("is a workbook-changing tool and is summarised for the change list", () => {
    expect(WRITE_TOOLS.has("record_test_results")).toBe(true);
    const change = describeChange("record_test_results", {}, { ok: true, data: { sheet: "Test Cases", rowsUpdated: 2, cellsWritten: 4, addedColumns: ["Tiến độ (E)"] } }, WRITE_TOOLS)!;
    expect(change.summary).toContain("Test Cases");
    expect(change.summary).toContain("Tiến độ (E)");
  });

  it("finds columns by header text in a header row that is not the first, and appends missing columns", async () => {
    const result = await record({ results: [
      { row: 4, values: { "Tiến độ": "Hoàn thành", "Kết quả thực tế": "Vào được trang chủ", "Kết quả": "Pass" } },
      { row: 5, values: { "tiến độ": "Hoàn thành", "Kết quả": "Fail" } },
    ] });
    expect(result).toMatchObject({ ok: true, data: { rowsUpdated: 2, cellsWritten: 5 } });
    expect((result as any).data.addedColumns).toEqual(["Tiến độ (E)", "Kết quả thực tế (F)", "Kết quả (G)"]);
    expect(book.get("Test Cases", "E3")).toBe("Tiến độ");
    expect(book.get("Test Cases", "F3")).toBe("Kết quả thực tế");
    expect(book.get("Test Cases", "E4")).toBe("Hoàn thành");
    expect(book.get("Test Cases", "E5")).toBe("Hoàn thành");
    expect(book.get("Test Cases", "G5")).toBe("Fail");
    // the test case itself is never touched
    expect(book.get("Test Cases", "C4")).toBe("Nhập user/pass hợp lệ");
    expect(book.get("Test Cases", "A1")).toBe("Project X - Test cases");
  });

  it("reuses a column that already exists, whatever its spacing and case", async () => {
    book.sheet("Test Cases").setCell(2, 4, { v: "  Kết  Quả ", f: null });
    const result = await record({ results: [{ row: 6, values: { "kết quả": "Pass" } }] });
    expect((result as any).data.addedColumns).toEqual([]);
    expect(book.get("Test Cases", "E6")).toBe("Pass");
  });

  it("can update its own cells (Đang test -> Pass) without asking, but refuses to overwrite anything else", async () => {
    await record({ results: [{ row: 4, values: { "Kết quả": "Đang test" } }] });
    expect(await record({ results: [{ row: 4, values: { "Kết quả": "Pass" } }] })).toMatchObject({ ok: true });
    expect(book.get("Test Cases", "E4")).toBe("Pass");

    book.sheet("Test Cases").setCell(4, 4, { v: "Người khác đã ghi", f: null });
    const refused = await record({ results: [{ row: 5, values: { "Kết quả": "Pass" } }, { row: 6, values: { "Kết quả": "Pass" } }] });
    expect(refused).toMatchObject({ ok: false, error: { code: "would_overwrite" } });
    expect(book.get("Test Cases", "E5")).toBe("Người khác đã ghi");
    expect(book.get("Test Cases", "E6")).toBeNull();
    expect(await record({ allowOverwrite: true, results: [{ row: 5, values: { "Kết quả": "Pass" } }] })).toMatchObject({ ok: true });
    expect(book.get("Test Cases", "E5")).toBe("Pass");
  });

  it("stores page text as text, never as a formula", async () => {
    await record({ results: [{ row: 4, values: { "Kết quả thực tế": "=HYPERLINK(\"http://x\")" } }] });
    expect(String(book.get("Test Cases", "E4"))).toBe("'=HYPERLINK(\"http://x\")");
  });

  it("refuses rows that are not data rows, a missing sheet, and new columns when told not to add them", async () => {
    expect(await record({ results: [{ row: 3, values: { A: "x" } }] })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await record({ results: [{ row: 99, values: { A: "x" } }] })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await record({ sheet: "Nope", results: [{ row: 4, values: { A: "x" } }] })).toMatchObject({ ok: false, error: { code: "sheet_not_found" } });
    expect(await record({ addMissingColumns: false, results: [{ row: 4, values: { "Tiến độ": "x" } }] })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(book.get("Test Cases", "E3")).toBeNull();
  });
});

describe("open_test_browser", () => {
  function fixture(store: TestSessions | null = null, knowledge: SiteKnowledge | null = null) {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("test-browser", null);
    session.capabilities.playwright = true;
    Object.defineProperty(session, "isConnected", { get: () => true, configurable: true });
    const gateway = new ToolGateway(sessions, { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 }, 1000,
      undefined, false, null, null, null, undefined, store, knowledge);
    const use = vi.spyOn(gateway.externalTools, "useProfile").mockResolvedValue();
    return { session, gateway, use, tool: (name: string, args: unknown) => gateway.invoke({ sessionId: session.id, tool: name, args }), call: (args: unknown) => gateway.invoke({ sessionId: session.id, tool: "open_test_browser", args }) };
  }

  it("normalises origins, only accepts plain http(s) sites, and is hidden until the browser tool is on", () => {
    expect(normalizeOrigin("https://staging.example.com/web#x")).toBe("https://staging.example.com");
    expect(normalizeOrigin("http://10.0.0.5:8069/odoo")).toBe("http://10.0.0.5:8069");
    for (const bad of ["file:///etc/passwd", "javascript:alert(1)", "https://u:p@example.com", "example.com", ""]) expect(normalizeOrigin(bad)).toBeNull();
    expect(() => parseToolInput("open_test_browser", { origins: ["ftp://x"] })).toThrow();
    expect(() => parseToolInput("open_test_browser", { origins: [] })).toThrow();
    expect(exposedToolNames(false, { playwright: true })).toEqual(expect.arrayContaining(["open_test_browser", "record_test_results"]));
    expect(exposedToolNames(false, {})).not.toContain("open_test_browser");
    expect(exposedToolNames(false, {})).toContain("record_test_results");
  });

  it("asks the user every time and then limits the browser to exactly the approved sites", async () => {
    const { session, use, call } = fixture();
    const ask = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: true } });
    const result = await call({ origins: ["https://staging.example.com/web/login", "https://staging.example.com", "https://sso.example.com"] });
    expect(ask).toHaveBeenCalledWith("__authorize_test_browser", { origins: ["https://staging.example.com", "https://sso.example.com"], visible: true }, 120_000);
    expect(use).toHaveBeenCalledWith(session.id, { allowedOrigins: ["https://staging.example.com", "https://sso.example.com"], visible: true });
    expect(result).toMatchObject({ ok: true, data: { ready: true, browserMayReachOnly: ["https://staging.example.com", "https://sso.example.com"] } });
  });

  it("does nothing when the user declines or the pane has gone", async () => {
    const { session, use, call } = fixture();
    const ask = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: false } });
    expect(await call({ origins: ["https://staging.example.com"] })).toMatchObject({ ok: false, error: { code: "cancelled" } });
    ask.mockResolvedValue({ ok: true, data: { approved: true } });
    Object.defineProperty(session, "isConnected", { get: () => false });
    expect(await call({ origins: ["https://staging.example.com"] })).toMatchObject({ ok: false, error: { code: "disconnected" } });
    expect(use).not.toHaveBeenCalled();
  });

  it("shows the user exactly which sites and whether the window is visible", () => {
    const text = describeTestBrowserApproval({ origins: ["https://a.example.com", "https://b.example.com"], visible: true });
    expect(text.target).toBe("https://a.example.com\nhttps://b.example.com");
    expect(text.detail).toContain("sign in");
    expect(describeTestBrowserApproval({ origins: ["https://a.example.com"], visible: false }).detail).toContain("hidden");
  });

  it("needs either a saved session or sites", () => {
    expect(() => parseToolInput("open_test_browser", {})).toThrow();
    expect(() => parseToolInput("open_test_browser", { session: "Admin Role" })).toThrow();
    expect(parseToolInput("open_test_browser", { session: "admin" })).toMatchObject({ session: "admin" });
  });

  describe("saved sessions (excel-partner login)", () => {
    const withStore = () => {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ep-sessions-"));
      const store = new TestSessions(dir);
      store.save("admin", { cookies: [{ name: "sid", value: "SECRET-COOKIE" }], origins: [] }, { origins: ["https://staging.example.com", "https://sso.example.com"], savedAt: "2026-10-29T10:00:00.000Z", url: "https://staging.example.com" });
      return { dir, store };
    };

    it("lists names, sites and dates but never the cookies", async () => {
      const { dir, store } = withStore();
      const { tool } = fixture(store);
      const result = await tool("list_test_sessions", {});
      expect(JSON.stringify(result)).not.toContain("SECRET-COOKIE");
      expect(result).toMatchObject({ ok: true, data: { sessions: [{ name: "admin", origins: ["https://staging.example.com", "https://sso.example.com"] }] } });
      expect(fs.statSync(path.join(dir, "admin.state.json")).isFile()).toBe(true);
    });

    it("starts the browser with the saved state and exactly the saved sites after approval", async () => {
      const { dir, store } = withStore();
      const { session, use, call } = fixture(store);
      const ask = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: true } });
      const result = await call({ session: "admin" });
      expect(ask).toHaveBeenCalledWith("__authorize_test_browser", { origins: ["https://staging.example.com", "https://sso.example.com"], visible: false, session: "admin" }, 120_000);
      expect(use).toHaveBeenCalledWith(session.id, { allowedOrigins: ["https://staging.example.com", "https://sso.example.com"], visible: false, storageState: path.join(dir, "admin.state.json") });
      expect(result).toMatchObject({ ok: true, data: { signedInAs: "admin" } });
      expect(JSON.stringify(result)).not.toContain("SECRET-COOKIE");
    });

    it("names the saved sessions when the requested one does not exist, without opening anything", async () => {
      const { store } = withStore();
      const { session, use, call } = fixture(store);
      const ask = vi.spyOn(session, "requestTool");
      const result = await call({ session: "nobody" });
      expect(result).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
      expect((result as any).error.message).toContain("admin");
      expect((result as any).error.message).toContain("excel-partner login");
      expect(ask).not.toHaveBeenCalled();
      expect(use).not.toHaveBeenCalled();
    });

    it("tells the user which saved sign-in the browser will use", () => {
      const text = describeTestBrowserApproval({ origins: ["https://a.example.com"], visible: false, session: "admin" });
      expect(text.detail).toContain('"admin"');
      expect(text.detail).toContain("never sees the cookies");
    });
  });
});

describe("saved knowledge about a system", () => {
  const origin = "https://staging.example.com";
  const setup = () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ep-knowledge-"));
    const knowledge = new SiteKnowledge(dir);
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("knowledge", null);
    session.capabilities.playwright = true;
    Object.defineProperty(session, "isConnected", { get: () => true, configurable: true });
    const gateway = new ToolGateway(sessions, { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 }, 1000, undefined, false, null, null, null, undefined, null, knowledge);
    const ask = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: true } });
    return { dir, knowledge, session, gateway, ask, tool: (name: string, args: unknown) => gateway.invoke({ sessionId: session.id, tool: name, args }) };
  };
  const note = { origin, name: "login", kind: "login", summary: "Odoo form login, then pick the company", content: "Open /web/login, use the saved session. The company switcher is top right; choose ACME." };

  it("keys notes by host and port, whatever the path or scheme details", () => {
    expect(siteKey("https://Staging.Example.com/web?x=1")).toBe("staging.example.com");
    expect(siteKey("http://localhost:8081/odoo")).toBe("localhost_8081");
    expect(siteKey("javascript:alert(1)")).toBeNull();
  });

  it("saves a note as a readable Markdown file and lists it with its summary and date", async () => {
    const { dir, tool } = setup();
    expect(await tool("save_site_knowledge", note)).toMatchObject({ ok: true, data: { saved: { name: "login", kind: "login" } } });
    const file = path.join(dir, "staging.example.com", "login.md");
    expect(fs.readFileSync(file, "utf8")).toContain("summary: Odoo form login, then pick the company");
    expect(await tool("read_site_knowledge", { origin })).toMatchObject({ ok: true, data: { notes: [{ name: "login", kind: "login", summary: "Odoo form login, then pick the company" }] } });
    expect(await tool("read_site_knowledge", { origin, name: "login" })).toMatchObject({ ok: true, data: { content: expect.stringContaining("company switcher") } });
  });

  it("asks the user first and saves nothing when they decline", async () => {
    const { dir, ask, tool } = setup();
    ask.mockResolvedValue({ ok: true, data: { approved: false } });
    expect(await tool("save_site_knowledge", note)).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(ask).toHaveBeenCalledWith("__authorize_site_knowledge", expect.objectContaining({ action: "save", name: "login", replaces: false }), 120_000);
    expect(fs.existsSync(path.join(dir, "staging.example.com"))).toBe(false);
  });

  it("refuses notes that hold passwords, tokens or cookies, before asking anyone", async () => {
    const { ask, tool } = setup();
    for (const content of ["password: Sup3rSecret!", "Authorization: Bearer abcdefghijklmnopqrstuvwxyz0123", "token=abcd1234efgh", "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.sig", "session_id: 8f2c9d"]) {
      expect(await tool("save_site_knowledge", { ...note, content })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    }
    expect(ask).not.toHaveBeenCalled();
  });

  it("replaces a note saved under the same name, and deletes one on request", async () => {
    const { ask, tool } = setup();
    await tool("save_site_knowledge", note);
    await tool("save_site_knowledge", { ...note, summary: "Updated", content: "New steps." });
    expect(ask).toHaveBeenLastCalledWith("__authorize_site_knowledge", expect.objectContaining({ replaces: true }), 120_000);
    expect(await tool("read_site_knowledge", { origin })).toMatchObject({ data: { notes: [{ summary: "Updated" }] } });
    expect(await tool("delete_site_knowledge", { origin, name: "login" })).toMatchObject({ ok: true, data: { deleted: true } });
    expect(await tool("read_site_knowledge", { origin })).toMatchObject({ data: { notes: [] } });
    expect(await tool("read_site_knowledge", { origin, name: "login" })).toMatchObject({ ok: false });
  });

  it("limits how much one system can accumulate", async () => {
    const { knowledge } = setup();
    for (let n = 0; n < 30; n++) knowledge.save(origin, { name: `note-${n}`, kind: "howto", summary: "s", content: "c" });
    expect(() => knowledge.save(origin, { name: "one-too-many", kind: "howto", summary: "s", content: "c" })).toThrow(/30 notes/);
    expect(() => knowledge.save(origin, { name: "note-1", kind: "howto", summary: "s", content: "x".repeat(70_000) })).toThrow(/exceed/);
  });

  it("hands the list of notes to the assistant when it opens the browser, and only for the approved sites", async () => {
    const { knowledge, gateway, session } = setup();
    knowledge.save(origin, { name: "login", kind: "login", summary: "How to get in", content: "Use the saved session." });
    knowledge.save("https://other.example.com", { name: "secret-place", kind: "quirk", summary: "Other site", content: "x" });
    vi.spyOn(gateway.externalTools, "useProfile").mockResolvedValue();
    const result: any = await gateway.invoke({ sessionId: session.id, tool: "open_test_browser", args: { origins: [origin, "https://new.example.com"] } });
    expect(result.data.knowledge.sites).toEqual([{ origin, notes: [expect.objectContaining({ name: "login", summary: "How to get in" })] }]);
    expect(result.data.knowledge.sitesWithoutNotes).toEqual(["https://new.example.com"]);
    expect(JSON.stringify(result)).not.toContain("secret-place");
    expect(JSON.stringify(result)).not.toContain("Use the saved session");
  });

  it("shows the whole note in the dialog so the user knows what will steer later sessions", () => {
    const text = describeSiteKnowledgeApproval({ action: "save", origin, name: "login", kind: "login", summary: "How to get in", content: "Step one.", replaces: false });
    expect(text.target).toBe("Step one.");
    expect(text.detail).toContain(origin);
    expect(describeSiteKnowledgeApproval({ action: "delete", origin, name: "login" }).title).toContain("Delete");
  });
});
