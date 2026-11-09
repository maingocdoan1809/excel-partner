import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { strToU8, zipSync } from "fflate";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { attachmentPath, createAttachment } from "../bridge/src/attachments.js";
import { ReferenceFiles } from "../bridge/src/reference-files.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { exposedToolNamesForHost, HOST_IDS, isToolNameForHost, parseToolInputForHost } from "../protocol/src/index.js";
import { sampleDeck, sampleDocx } from "./support/office-files.js";

let root: string;
beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), "docfiles-"));
  fs.writeFileSync(path.join(root, "report.docx"), sampleDocx());
  fs.writeFileSync(path.join(root, "deck.pptx"), sampleDeck());
  fs.writeFileSync(path.join(root, "notes.txt"), "plain text");
});
afterEach(() => fs.rmSync(root, { recursive: true, force: true }));

function setup(host: "excel" | "word" | "powerpoint", roots = [root]) {
  const sessions = new SessionManager({ ttlMs: 60_000 });
  const session = sessions.create(`conn-${host}-0000000`, { host });
  const limits = { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 };
  const gateway = new ToolGateway(sessions, limits, 1000, undefined, false, null, new ReferenceFiles(roots));
  const read = (args: Record<string, unknown>) => gateway.invoke({ sessionId: session.id, tool: "read_document_file", args });
  return { session, gateway, read };
}

describe("read_document_file", () => {
  it("is offered in every Office app, together with the Excel file reader", () => {
    for (const host of HOST_IDS) {
      expect(isToolNameForHost(host, "read_document_file")).toBe(true);
      expect(isToolNameForHost(host, "read_workbook_file")).toBe(true);
      expect(exposedToolNamesForHost(host, true)).toEqual(expect.arrayContaining(["read_document_file", "read_workbook_file", "list_reference_files"]));
    }
    expect(isToolNameForHost("word", "insert_sheets_from_file")).toBe(false);
    expect(parseToolInputForHost("word", "read_document_file", { path: "a.docx" })).toMatchObject({ from: 1, maxChars: 20_000 });
  });

  it.each(["excel", "word", "powerpoint"] as const)("reads a document and a deck from an allowed folder in a %s session", async (host) => {
    const { read } = setup(host);
    const docx = await read({ path: path.join(root, "report.docx") });
    expect(docx).toMatchObject({ ok: true, data: { format: "docx" } });
    expect(docx.ok && (docx.data as any).blocks[0]).toMatchObject({ text: "Annual report", level: 1 });
    const deck = await read({ path: path.join(root, "deck.pptx"), from: 1, count: 1 });
    expect(deck).toMatchObject({ ok: true, data: { format: "pptx", slideCount: 3, slides: [{ title: "Quarterly Review", notes: "Say thank you first." }] } });
  });

  it("reads an attachment of the same chat only", async () => {
    const { session, read } = setup("word", []);
    const other = setup("word", []);
    const attached = createAttachment("brief.docx", sampleDocx());
    session.attachments.set(attached.id, attached);
    expect(await read({ path: attachmentPath(attached.id) })).toMatchObject({ ok: true, data: { name: "brief.docx", format: "docx" } });
    expect(await other.read({ path: attachmentPath(attached.id) })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("asks for access to a file outside the allowed folders, once, and not for the wrong type", async () => {
    const { session, read } = setup("powerpoint", []);
    Object.defineProperty(session, "isConnected", { get: () => true });
    const ask = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: true } });
    const target = path.join(root, "deck.pptx");
    expect(await read({ path: target })).toMatchObject({ ok: true });
    expect(ask).toHaveBeenCalledWith("__authorize_reference_path", { path: target, resolvedPath: target }, 120_000);
    expect((await read({ path: target })).ok).toBe(true);
    expect(ask).toHaveBeenCalledTimes(1);
    ask.mockClear();
    expect(await read({ path: path.join(root, "notes.txt") })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(ask).not.toHaveBeenCalled();
  });

  it("respects a refusal", async () => {
    const { session, read } = setup("word", []);
    vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: false } });
    expect(await read({ path: path.join(root, "report.docx") })).toMatchObject({ ok: false, error: { code: "cancelled" } });
  });

  it("answers a broken file with a message, not a crash", async () => {
    fs.writeFileSync(path.join(root, "broken.docx"), "this is not a zip");
    const { read } = setup("word");
    expect(await read({ path: path.join(root, "broken.docx") })).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("not a valid Office document") } });
    expect(await read({ path: path.join(root, "missing.docx") })).toMatchObject({ ok: false });
  });

  it("lets a Word or PowerPoint session read an Excel file from disk too", async () => {
    const sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
    const workbook = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Main" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>';
    fs.writeFileSync(path.join(root, "book.xlsx"), zipSync({ "xl/workbook.xml": strToU8(workbook), "xl/_rels/workbook.xml.rels": strToU8(rels), "xl/worksheets/sheet1.xml": strToU8(sheet) }));
    for (const host of ["word", "powerpoint"] as const) {
      const { gateway, session } = setup(host);
      const result = await gateway.invoke({ sessionId: session.id, tool: "read_workbook_file", args: { path: path.join(root, "book.xlsx") } });
      expect(result.ok && (result.data as { sheets: { name: string }[] }).sheets[0]!.name).toBe("Main");
    }
  });
});
