import { describe, expect, it } from "vitest";
import type { WebSocket } from "ws";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { OfficeRuntime } from "../bridge/src/office-runtime.js";
import type { WorkbookSession } from "../bridge/src/session.js";
import { CONNECTED_READ_TOOLS_BY_HOST, ConnectedFileSchema, WORKBOOK_NAME_REQUEST, parseToolInput, type HostId, type ToolResult } from "../protocol/src/index.js";

const limits = { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 };

interface Sent { tool: string; args: unknown }

/** A pane that answers every tool_request itself: approvals with `approve`, reads with `reply`. */
function attachPane(session: WorkbookSession, options: { approve?: boolean; reply?: (tool: string) => ToolResult } = {}) {
  const sent: Sent[] = [];
  session.attachSocket({
    readyState: 1,
    close() {},
    send(raw: string) {
      const message = JSON.parse(raw) as { type: string; callId?: string; tool?: string; args?: unknown };
      if (message.type !== "tool_request") return;
      if (message.tool === WORKBOOK_NAME_REQUEST) { queueMicrotask(() => session.resolveTool(message.callId!, { ok: true, data: { name: session.workbookName } })); return; }
      sent.push({ tool: message.tool!, args: message.args });
      const result: ToolResult = message.tool!.startsWith("__authorize_")
        ? { ok: true, data: { approved: options.approve ?? true } }
        : options.reply?.(message.tool!) ?? { ok: true, data: { from: session.fileId } };
      queueMicrotask(() => session.resolveTool(message.callId!, result));
    },
  } as unknown as WebSocket);
  // Another file is open (these tests are about work across files), so the tools for it are offered, as they are to a real session.
  session.capabilities = { ...session.capabilities, connectedFiles: true };
  return sent;
}

/** One open file per Office app, each with its own pane. */
function office() {
  const sessions = new WorkbookSessions({ ttlMs: 60_000 });
  const made = {} as Record<HostId, { session: WorkbookSession; sent: Sent[] }>;
  const names: Record<HostId, string> = { excel: "Budget.xlsx", word: "Plan.docx", powerpoint: "Review.pptx" };
  const data: Record<HostId, unknown> = {
    excel: { sheets: ["Q1"] },
    word: { paragraphs: [{ index: 0, text: "Revenue target is 12M" }] },
    powerpoint: { slides: [{ id: "256", title: "Quarterly Review" }] },
  };
  for (const host of ["excel", "word", "powerpoint"] as const) {
    const session = sessions.create(`conn-${host}-12345678`, { host });
    session.workbookName = names[host];
    made[host] = { session, sent: attachPane(session, { reply: () => ({ ok: true, data: data[host] }) }) };
  }
  const gateway = new OfficeRuntime(sessions, limits, 1000);
  const read = (from: HostId, file: string, tool: string, args: unknown = {}) =>
    gateway.invoke({ sessionId: made[from].session.id, tool: "read_connected_file", args: { file, tool, arguments: args } });
  return { sessions, gateway, made, read };
}

describe("files of different Office apps", () => {
  it("lists every open file with the application it belongs to, in every app, and says which tools read each kind", async () => {
    const { gateway, made } = office();
    for (const host of ["excel", "word", "powerpoint"] as const) {
      const result = await gateway.invoke({ sessionId: made[host].session.id, tool: "list_connected_files", args: {} });
      expect(result.ok).toBe(true);
      const data = (result as { data: { files: Array<{ host: string; current: boolean; name: string }>; readTools: Record<string, string[]> } }).data;
      expect(data.files).toHaveLength(3);
      expect(data.files[0]).toMatchObject({ host, current: true });
      expect(data.files.map((file) => file.host).sort()).toEqual(["excel", "powerpoint", "word"]);
      expect(Object.keys(data.readTools).sort()).toEqual(["excel", "powerpoint", "word"].filter((other) => other !== host));
    }
  });

  it("lets a Word document read a workbook and a deck, each with that application's own tool, after the user allows it", async () => {
    const { made, read } = office();
    made.word.session.beginTurn("req-1");
    const sheet = await read("word", made.excel.session.fileId, "get_sheets");
    expect(sheet).toMatchObject({ ok: true, data: { file: { name: "Budget.xlsx", host: "excel" }, tool: "get_sheets", result: { sheets: ["Q1"] } } });
    const deck = await read("word", made.powerpoint.session.fileId, "read_slide_text", { ref: "256:2" });
    expect(deck).toMatchObject({ ok: true, data: { file: { name: "Review.pptx", host: "powerpoint" }, tool: "read_slide_text", result: { slides: [{ title: "Quarterly Review" }] } } });
    expect(made.word.sent.map((item) => item.tool)).toEqual(["__authorize_connected_file", "__authorize_connected_file"]); // one question per file
    await read("word", made.excel.session.fileId, "get_sheets");
    expect(made.word.sent).toHaveLength(2); // the same file again in the same message is not asked again
    expect(made.excel.sent.map((item) => item.tool)).toEqual(["get_sheets", "get_sheets"]);
    expect(made.powerpoint.sent.map((item) => item.tool)).toEqual(["read_slide_text"]);
  });

  it("tells the person what kind of file is about to be read", async () => {
    const { made, read } = office();
    made.excel.session.beginTurn("req-1");
    await read("excel", made.word.session.fileId, "read_doc_section", { count: 3 });
    expect(made.excel.sent[0]).toEqual({ tool: "__authorize_connected_file", args: expect.objectContaining({ name: "Plan.docx", host: "word" }) });
  });

  it("lets a workbook or a deck read a Word document, and gives no cell link for a file that has no cells", async () => {
    const { made, read } = office();
    made.excel.session.beginTurn("req-1");
    const fromExcel = await read("excel", made.word.session.fileId, "read_doc_section", { count: 5 });
    expect(fromExcel).toMatchObject({ ok: true, data: { file: { host: "word" }, result: { paragraphs: [{ text: "Revenue target is 12M" }] } } });
    expect((fromExcel as { data: Record<string, unknown> }).data.link).toBeUndefined();
    made.powerpoint.session.beginTurn("req-2");
    expect(await read("powerpoint", made.word.session.fileId, "search_doc_text", { query: "target" })).toMatchObject({ ok: true });
  });

  it("refuses a tool that belongs to another application, and never asks the person or the other pane", async () => {
    const { made, read } = office();
    made.word.session.beginTurn("req-1");
    const wrong = await read("word", made.excel.session.fileId, "read_doc_section");
    expect(wrong).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("cannot read a Excel file") } });
    expect(wrong.ok === false && wrong.error.message).toContain(CONNECTED_READ_TOOLS_BY_HOST.excel.join(", "));
    expect(await read("word", made.powerpoint.session.fileId, "read_range", { address: "A1" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(made.word.sent).toEqual([]);
    expect(made.excel.sent).toEqual([]);
    expect(made.powerpoint.sent).toEqual([]);
  });

  it("never reaches a write tool of another file, and checks the arguments for that file's application", async () => {
    const { made, read } = office();
    made.excel.session.beginTurn("req-1");
    expect(await read("excel", made.word.session.fileId, "edit_doc_text", { old_text: "a", new_text: "b" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(() => parseToolInput("read_connected_file", { file: "f2", tool: "edit_doc_text" })).toThrow();
    expect(await read("excel", made.word.session.fileId, "read_doc_section", { count: 0 })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(made.word.sent).toEqual([]);
  });

  it("declines when the person says no, for any kind of file", async () => {
    const sessions = new WorkbookSessions({ ttlMs: 60_000 });
    const doc = sessions.create("conn-word-87654321", { host: "word" });
    const deck = sessions.create("conn-ppt-87654321", { host: "powerpoint" });
    doc.workbookName = "Plan.docx"; deck.workbookName = "Review.pptx";
    attachPane(doc, { approve: false });
    const deckSent = attachPane(deck);
    const gateway = new OfficeRuntime(sessions, limits, 1000);
    doc.beginTurn("req-1");
    const result = await gateway.invoke({ sessionId: doc.id, tool: "read_connected_file", args: { file: deck.fileId, tool: "read_slide_text", arguments: { ref: "256:2" } } });
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled", message: expect.stringContaining("Review.pptx") } });
    expect(deckSent).toEqual([]);
  });

  it("names an unnamed file after its kind", async () => {
    const { made, gateway } = office();
    made.word.session.workbookName = null;
    made.powerpoint.session.workbookName = null;
    const result = await gateway.invoke({ sessionId: made.excel.session.id, tool: "list_connected_files", args: {} });
    const names = (result as { data: { files: Array<{ host: string; name: string }> } }).data.files.map((file) => `${file.host}=${file.name}`).sort();
    expect(names).toEqual(["excel=Budget.xlsx", "powerpoint=Untitled presentation", "word=Untitled document"]);
  });

  it("keeps working with a pane that does not send a host: such a file is a workbook", () => {
    expect(ConnectedFileSchema.parse({ id: "f1", name: "Old.xlsx", current: false })).toMatchObject({ host: "excel" });
  });
});
