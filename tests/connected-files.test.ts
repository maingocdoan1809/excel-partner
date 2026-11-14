import { describe, expect, it } from "vitest";
import type { WebSocket } from "ws";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { OfficeRuntime } from "../bridge/src/office-runtime.js";
import type { WorkbookSession } from "../bridge/src/session.js";
import { ClientMessageSchema, MAX_DELEGATED_TEXT, ServerMessageSchema, UNNAMED_WORKBOOK, WORKBOOK_NAME_REQUEST, parseToolInput, type ToolResult } from "../protocol/src/index.js";

const limits = { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 };

interface Sent { tool: string; args: unknown; requestId: string }

/** A pane that answers every tool_request itself: approvals with `approve`, reads with `reply`. */
function attachPane(session: WorkbookSession, options: { approve?: boolean; reply?: (tool: string, args: unknown) => ToolResult } = {}) {
  const sent: Sent[] = [];
  const socket = {
    readyState: 1,
    close() {},
    send(raw: string) {
      const message = JSON.parse(raw) as { type: string; callId?: string; tool?: string; args?: unknown; requestId?: string };
      if (message.type !== "tool_request") return;
      // The bridge asking what the workbook is called is not a tool of the chat: answer it without recording it.
      if (message.tool === WORKBOOK_NAME_REQUEST) {
        queueMicrotask(() => session.resolveTool(message.callId!, { ok: true, data: { name: session.workbookName } }));
        return;
      }
      sent.push({ tool: message.tool!, args: message.args, requestId: message.requestId! });
      const result: ToolResult = message.tool!.startsWith("__authorize_")
        ? { ok: true, data: { approved: options.approve ?? true } }
        : options.reply?.(message.tool!, message.args) ?? { ok: true, data: { from: session.fileId } };
      queueMicrotask(() => session.resolveTool(message.callId!, result));
    },
  } as unknown as WebSocket;
  session.attachSocket(socket);
  // Another file is open (these tests are about work across files), so the tools for it are offered, as they are to a real session.
  session.capabilities = { ...session.capabilities, connectedFiles: true };
  return sent;
}

function setup(options: { approve?: boolean } = {}) {
  const sessions = new WorkbookSessions({ ttlMs: 60_000 });
  const here = sessions.create("conn-here", null);
  const there = sessions.create("conn-there", null);
  here.workbookName = "Budget 2026.xlsx";
  there.workbookName = "Actuals.xlsx";
  const hereSent = attachPane(here, { approve: options.approve });
  const thereSent = attachPane(there, { reply: (tool) => ({ ok: true, data: { tool, sheets: ["Q1"] } }) });
  const gateway = new OfficeRuntime(sessions, limits, 1000);
  return { sessions, here, there, hereSent, thereSent, gateway };
}

const read = (gateway: OfficeRuntime, session: WorkbookSession, file: string, tool = "get_sheets", args: unknown = {}) =>
  gateway.invoke({ sessionId: session.id, tool: "read_connected_file", args: { file, tool, arguments: args } });

describe("connected files", () => {
  it("lists every connected workbook with the current one first", async () => {
    const { gateway, there } = setup();
    const result = await gateway.invoke({ sessionId: there.id, tool: "list_connected_files", args: {} });
    expect(result).toMatchObject({ ok: true, data: { files: [
      { id: "f2", name: "Actuals.xlsx", host: "excel", current: true },
      { id: "f1", name: "Budget 2026.xlsx", host: "excel", current: false },
    ] } });
  });

  it("leaves out a workbook whose task pane has gone", async () => {
    const { gateway, here, there } = setup();
    there.detachSocket(there.socket!);
    const result = await gateway.invoke({ sessionId: here.id, tool: "list_connected_files", args: {} });
    expect((result as { data: { files: unknown[] } }).data.files).toHaveLength(1);
  });

  it("reads from the other workbook after the user allows it, and only that workbook's pane runs the read", async () => {
    const { gateway, here, there, hereSent, thereSent } = setup();
    here.beginTurn("req-1");
    const result = await read(gateway, here, there.fileId, "read_range", { sheet: { name: "Q1" }, address: "A1:B2" });
    expect(result).toMatchObject({ ok: true, data: { file: { id: "f2", name: "Actuals.xlsx" }, tool: "read_range", result: { tool: "read_range" } } });
    expect(hereSent.map((item) => item.tool)).toEqual(["__authorize_connected_file"]);
    expect(thereSent).toEqual([{ tool: "read_range", args: expect.objectContaining({ address: "A1:B2" }), requestId: "no-turn" }]);
  });

  it("hands back the range it read as a link ready to paste, never one that could point somewhere else", async () => {
    const { gateway, here, there } = setup();
    here.beginTurn("req-1");
    const link = async (tool: string, args: unknown) => ((await read(gateway, here, there.fileId, tool, args)) as { data: Record<string, unknown> }).data.link;
    expect(await link("read_range", { sheet: { name: "Q1" }, address: "A1:B2" })).toBe("[A1:B2](cell:?file=Actuals.xlsx&sheet=Q1&range=A1:B2)");
    expect(await link("read_range", { sheet: { name: "Sales 2026" }, address: "$A$1" })).toBe("[A1](cell:?file=Actuals.xlsx&sheet=Sales%202026&range=A1)");
    // No sheet asked for: the file's open sheet was read, and the link opens that sheet too.
    expect(await link("read_range", { address: "C3" })).toBe("[C3](cell:?file=Actuals.xlsx&range=C3)");
    // Nothing to name, or a sheet that only has an id: no link rather than a wrong one.
    expect(await link("get_sheets", {})).toBeUndefined();
    expect(await link("read_range", { sheet: { id: "{123}" }, address: "A1" })).toBeUndefined();
    // A range that is not one is refused before anything is read, so there is nothing to link.
    expect(await read(gateway, here, there.fileId, "read_range", { sheet: { name: "Q1" }, address: "not a range" })).toMatchObject({ ok: false });
  });

  it("asks once per message, then again for the next message", async () => {
    const { gateway, here, there, hereSent } = setup();
    here.beginTurn("req-1");
    await read(gateway, here, there.fileId);
    await read(gateway, here, there.fileId);
    expect(hereSent.filter((item) => item.tool === "__authorize_connected_file")).toHaveLength(1);
    here.endTurn("req-1");
    here.beginTurn("req-2");
    await read(gateway, here, there.fileId);
    expect(hereSent.filter((item) => item.tool === "__authorize_connected_file")).toHaveLength(2);
  });

  it("does not touch the other workbook when the user declines", async () => {
    const { gateway, here, there, thereSent } = setup({ approve: false });
    here.beginTurn("req-1");
    const result = await read(gateway, here, there.fileId);
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(thereSent).toEqual([]);
  });

  it("refuses anything that is not a read tool, so nothing can be written to another file", async () => {
    const { gateway, here, there, thereSent } = setup();
    here.beginTurn("req-1");
    for (const tool of ["write_range", "set_formula", "clear_range", "manage_sheet", "execute_office_js", "run_python"]) {
      expect(await read(gateway, here, there.fileId, tool, { address: "A1" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    }
    expect(thereSent).toEqual([]);
  });

  it("rejects the current workbook, an unknown id and bad arguments", async () => {
    const { gateway, here, there, thereSent } = setup();
    here.beginTurn("req-1");
    expect(await read(gateway, here, here.fileId)).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await read(gateway, here, "f99")).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await read(gateway, here, there.fileId, "read_range", { address: "1A" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(thereSent).toEqual([]);
  });

  it("applies the usual read size limit to the other workbook", async () => {
    const { gateway, here, there, thereSent } = setup();
    here.beginTurn("req-1");
    expect(await read(gateway, here, there.fileId, "read_range", { address: "A1:Z1000" })).toMatchObject({ ok: false, error: { code: "size_limit_exceeded" } });
    expect(thereSent).toEqual([]);
  });

  it("names the file in an error from the other workbook", async () => {
    const sessions = new WorkbookSessions({ ttlMs: 60_000 });
    const here = sessions.create("a", null);
    const there = sessions.create("b", null);
    there.workbookName = "Actuals.xlsx";
    attachPane(here);
    attachPane(there, { reply: () => ({ ok: false, error: { code: "invalid_arguments", message: "sheet 'Q9' not found" } }) });
    here.beginTurn("req-1");
    const result = await read(new OfficeRuntime(sessions, limits, 1000), here, there.fileId, "get_sheets");
    expect(result).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: "[Actuals.xlsx] sheet 'Q9' not found" } });
  });

  it("treats a workbook whose pane has gone as unknown", async () => {
    const { gateway, here, there } = setup();
    here.beginTurn("req-1");
    there.detachSocket(there.socket!);
    expect(await read(gateway, here, there.fileId)).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });
});

describe("connected files messages", () => {
  it("accepts workbook_info from the pane and connected_files from the bridge", () => {
    expect(ClientMessageSchema.parse({ type: "workbook_info", name: " Budget.xlsx " })).toEqual({ type: "workbook_info", name: "Budget.xlsx" });
    expect(() => ClientMessageSchema.parse({ type: "workbook_info", name: "  " })).toThrow();
    const files = [{ id: "f1", name: "Budget.xlsx", host: "excel", current: true }];
    expect(ServerMessageSchema.parse({ type: "connected_files", files })).toEqual({ type: "connected_files", files });
  });

  it("defaults the other tool's arguments and only names read tools", () => {
    expect(parseToolInput("read_connected_file", { file: "f2", tool: "get_sheets" }).arguments).toEqual({});
    expect(() => parseToolInput("read_connected_file", { file: "f2", tool: "write_range" })).toThrow();
  });
});

describe("connected files: review fixes", () => {
  it("rejects arguments passed next to `tool` instead of silently dropping them", () => {
    expect(() => parseToolInput("read_connected_file", { file: "f2", tool: "find_errors", sheet: "Data" })).toThrow();
    expect(parseToolInput("read_connected_file", { file: "f2", tool: "find_errors", arguments: { sheet: "Data" } }).arguments).toEqual({ sheet: "Data" });
    expect(() => parseToolInput("send_to_connected_file", { file: "f2", message: "x", extra: 1 })).toThrow();
  });

  it("shares one limit between the tool input and the delegated_prompt message", () => {
    expect(() => parseToolInput("send_to_connected_file", { file: "f2", message: "a".repeat(MAX_DELEGATED_TEXT) })).not.toThrow();
    expect(() => parseToolInput("send_to_connected_file", { file: "f2", message: "a".repeat(MAX_DELEGATED_TEXT + 1) })).toThrow();
    const prompt = { type: "delegated_prompt", delegationId: "d1", from: { id: "f1", name: "A" } };
    expect(() => ServerMessageSchema.parse({ ...prompt, text: "a".repeat(MAX_DELEGATED_TEXT) })).not.toThrow();
    expect(() => ServerMessageSchema.parse({ ...prompt, text: "a".repeat(MAX_DELEGATED_TEXT + 1) })).toThrow();
  });

  it("gives the assistant and the task pane the same list, and the assistant a name for an unnamed file", async () => {
    const { sessions, gateway, here, there } = setup();
    there.workbookName = null;
    // What the pane is sent: the unknown name stays empty, and the pane shows UNNAMED_WORKBOOK for it.
    expect(sessions.connectedFiles(here)).toEqual([
      { id: "f1", name: "Budget 2026.xlsx", host: "excel", current: true },
      { id: "f2", name: "", host: "excel", current: false },
    ]);
    const listed = await gateway.invoke({ sessionId: here.id, tool: "list_connected_files", args: {} });
    expect((listed as { data: { files: unknown[] } }).data.files).toEqual([
      { id: "f1", name: "Budget 2026.xlsx", host: "excel", current: true },
      { id: "f2", name: UNNAMED_WORKBOOK, host: "excel", current: false },
    ]);
  });

  it("blocks the same read the way the normal path does after Excel dropped in the middle of it", async () => {
    const { gateway, here, there } = setup();
    // The other file's pane goes away while it is running the read, which is how a crashed Excel looks from here.
    there.detachSocket(there.socket!);
    const dropping = { readyState: 1, close() {}, send(raw: string) {
      const message = JSON.parse(raw) as { tool?: string; callId?: string };
      if (message.tool === WORKBOOK_NAME_REQUEST) { queueMicrotask(() => there.resolveTool(message.callId!, { ok: true, data: { name: there.workbookName } })); return; }
      there.detachSocket(dropping);
    } } as unknown as WebSocket;
    there.attachSocket(dropping);
    here.beginTurn("req-1");
    expect(await read(gateway, here, "f2", "get_sheets", {})).toMatchObject({ ok: false, error: { code: "disconnected" } });
    // The identical call, made directly, is now refused; it shares the memory of crashed calls.
    expect(await gateway.invoke({ sessionId: here.id, tool: "get_sheets", args: {} })).toMatchObject({ ok: false, error: { code: "cancelled", message: expect.stringContaining("Blocked") } });
  });

  it("drops a read that the asker no longer waits for, so the other pane's own calls are not held up", async () => {
    const { there } = setup();
    const slow = { readyState: 1, close() {}, send() { /* never answers */ } } as unknown as WebSocket;
    there.detachSocket(there.socket!);
    there.attachSocket(slow);
    const controller = new AbortController();
    const pending = there.requestPeerRead("get_sheets", {}, 5000, controller.signal);
    controller.abort();
    expect(await pending).toMatchObject({ ok: false, error: { code: "cancelled" } });
    // Nothing is left pending, and the pane's queue moves on to the next call.
    expect(there.pendingToolCount).toBe(0);
  });
});

describe("reading a workbook's name when it is needed", () => {
  function pane(session: WorkbookSession, answer: () => ToolResult | null) {
    session.attachSocket({ readyState: 1, close() {}, send(raw: string) {
      const message = JSON.parse(raw) as { tool: string; callId: string };
      const result = message.tool === WORKBOOK_NAME_REQUEST ? answer() : { ok: true, data: {} } as ToolResult;
      if (result) queueMicrotask(() => session.resolveTool(message.callId, result));
    } } as unknown as WebSocket);
  }

  it("takes the new name and says it changed, then says nothing changed", async () => {
    const sessions = new WorkbookSessions({ ttlMs: 60_000 });
    const file = sessions.create("a", null);
    file.workbookName = "Budget.xlsx";
    pane(file, () => ({ ok: true, data: { name: "  Budget (final).xlsx " } }));
    expect(await file.refreshWorkbookName()).toBe(true);
    expect(file.workbookName).toBe("Budget (final).xlsx");
    expect(await file.refreshWorkbookName()).toBe(false);
  });

  it("keeps the name it has when the pane gives none, gives nonsense, or does not answer in time", async () => {
    const sessions = new WorkbookSessions({ ttlMs: 60_000 });
    const file = sessions.create("a", null);
    file.workbookName = "Kept.xlsx";
    for (const data of [{ name: null }, { name: "   " }, { name: 42 }, {}, null]) {
      file.detachSocket(file.socket!);
      pane(file, () => ({ ok: true, data }));
      expect(await file.refreshWorkbookName()).toBe(false);
    }
    file.detachSocket(file.socket!);
    pane(file, () => null);
    expect(await file.refreshWorkbookName(30)).toBe(false);
    expect(file.workbookName).toBe("Kept.xlsx");
    file.detachSocket(file.socket!);
    expect(await file.refreshWorkbookName()).toBe(false);
  });

  it("does not wait behind a long call of that pane", async () => {
    const sessions = new WorkbookSessions({ ttlMs: 60_000 });
    const file = sessions.create("a", null);
    file.workbookName = "Busy.xlsx";
    pane(file, () => ({ ok: true, data: { name: "Busy (renamed).xlsx" } }));
    const sent: string[] = [];
    const answering = file.socket as unknown as { send(raw: string): void };
    const original = answering.send.bind(answering);
    answering.send = (raw) => { sent.push(JSON.parse(raw).tool); if (JSON.parse(raw).tool === WORKBOOK_NAME_REQUEST) original(raw); };
    file.beginTurn("r1");
    void file.requestTool("read_range", {}, 5000); // a call that never gets an answer
    expect(await file.refreshWorkbookName()).toBe(true);
    expect(sent).toContain(WORKBOOK_NAME_REQUEST);
  });

  it("lists a renamed file under its new name when the assistant asks", async () => {
    const { gateway, here, there } = setup();
    there.detachSocket(there.socket!);
    pane(there, () => ({ ok: true, data: { name: "Actuals (final).xlsx" } }));
    const listed = await gateway.invoke({ sessionId: here.id, tool: "list_connected_files", args: {} });
    expect((listed as { data: { files: unknown[] } }).data.files).toEqual([
      { id: "f1", name: "Budget 2026.xlsx", host: "excel", current: true },
      { id: "f2", name: "Actuals (final).xlsx", host: "excel", current: false },
    ]);
  });

  it("accepts refresh_connected_files from a pane", () => {
    expect(ClientMessageSchema.parse({ type: "refresh_connected_files" })).toEqual({ type: "refresh_connected_files" });
  });
});
