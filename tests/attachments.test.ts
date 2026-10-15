import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { AttachmentError, attachmentPath, createAttachment } from "../bridge/src/attachments.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";

const png = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0]);
const xlsx = Buffer.from(zipSync({
  "xl/workbook.xml": strToU8('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Main" sheetId="1" r:id="rId1"/></sheets></workbook>'),
  "xl/_rels/workbook.xml.rels": strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'),
  "xl/worksheets/sheet1.xml": strToU8('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'),
}));

describe("user-selected attachments", () => {
  it("accepts general files and strips path-like names", () => {
    expect(createAttachment("C:\\Users\\me\\photo.png", png)).toMatchObject({ name: "photo.png", kind: "image", mimeType: "image/png" });
    expect(createAttachment("book.xlsx", xlsx)).toMatchObject({ name: "book.xlsx", kind: "workbook" });
    expect(createAttachment("notes.md", Buffer.from("# Hello"))).toMatchObject({ name: "notes.md", kind: "file" });
    expect(createAttachment("photo.png", Buffer.from("not a PNG"))).toMatchObject({ kind: "file" });
    expect(() => createAttachment("empty.md", Buffer.alloc(0))).toThrow(AttachmentError);
  });

  it("lets the CLI read text and individual ZIP entries from multiple attachments", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("0123456789abcdef", null);
    const md = createAttachment("notes.md", Buffer.from("# Notes\nFirst file"));
    const zip = createAttachment("bundle.zip", Buffer.from(zipSync({ "readme.txt": strToU8("Second file") })));
    session.attachments.set(md.id, md);
    session.attachments.set(zip.id, zip);
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000);
    const read = (path: string, entry?: string) => gateway.invoke({ sessionId: session.id, tool: "read_attached_file", args: { path, ...(entry ? { entry } : {}) } });
    expect(await read(attachmentPath(md.id))).toMatchObject({ ok: true, data: { encoding: "utf-8", content: "# Notes\nFirst file" } });
    expect(await read(attachmentPath(zip.id))).toMatchObject({ ok: true, data: { format: "zip", entries: [{ name: "readme.txt" }] } });
    expect(await read(attachmentPath(zip.id), "readme.txt")).toMatchObject({ ok: true, data: { content: "Second file" } });
    expect(await read(attachmentPath(zip.id), "missing.txt")).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("serves an uploaded image and workbook only in their workbook session", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const one = sessions.create("0123456789abcdef", null);
    const two = sessions.create("fedcba9876543210", null);
    const image = createAttachment("photo.png", png);
    const book = createAttachment("book.xlsx", xlsx);
    one.attachments.set(image.id, image);
    one.attachments.set(book.id, book);
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000);
    const view = await gateway.invoke({ sessionId: one.id, tool: "read_attached_image", args: { path: attachmentPath(image.id) } });
    expect(view).toMatchObject({ ok: true, data: { image: { mimeType: "image/png", data: png.toString("base64") } } });
    const read = await gateway.invoke({ sessionId: one.id, tool: "read_workbook_file", args: { path: attachmentPath(book.id) } });
    expect(read.ok).toBe(true);
    if (read.ok) expect((read.data as { sheets: { name: string }[] }).sheets[0]!.name).toBe("Main");
    expect(await gateway.invoke({ sessionId: two.id, tool: "read_attached_image", args: { path: attachmentPath(image.id) } })).toMatchObject({ ok: false });
    expect(await gateway.invoke({ sessionId: two.id, tool: "read_workbook_file", args: { path: attachmentPath(book.id) } })).toMatchObject({ ok: false });
  });
});
