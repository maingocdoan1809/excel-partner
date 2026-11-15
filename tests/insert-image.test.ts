import fs from "node:fs";
import http from "node:http";
import type { AddressInfo } from "node:net";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createAttachment } from "../bridge/src/attachments.js";
import { detectImage } from "../bridge/src/image-info.js";
import { ImageDownloadError, downloadImage, isPrivateAddress } from "../bridge/src/image-download.js";
import { OfficeRuntime } from "../bridge/src/office-runtime.js";
import { ReferenceFiles } from "../bridge/src/reference-files.js";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { fitImage, naturalSize, centreIn } from "../frontend/taskpane/executor/image-layout.js";
import { executeSharedTool } from "../frontend/taskpane/host/shared.js";
import { HOST_IDS, InsertImageInputSchema, exposedToolNamesForHost, SHARED_WRITE_TOOLS, type ResolvedInsertImageInput } from "../protocol/src/index.js";
import { ALL_TOOL_DEFINITIONS } from "../shared/src/tool-definitions.js";
import { definitionForHost } from "../shared/src/host-overrides.js";
import { FakeWordDoc, installWordFake, uninstallWordFake } from "./support/fake-word.js";
import { MockWorkbook, installOfficeMock, type MockSheet } from "./support/office-mock.js";
import { useFakeDeck } from "./support/ppt-fixture.js";

/** A PNG header with the given size: enough for the bridge, which only reads magic bytes and the size. */
function png(width: number, height: number): Buffer {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(13);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), length, Buffer.from("IHDR"), header, Buffer.alloc(4), Buffer.from("IEND")]);
}

function jpeg(width: number, height: number): Buffer {
  const size = Buffer.alloc(4);
  size.writeUInt16BE(height, 0);
  size.writeUInt16BE(width, 2);
  return Buffer.concat([
    Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00]),
    Buffer.from([0xff, 0xc0, 0x00, 0x0b, 0x08]), size, Buffer.from([0x01, 0x01, 0x11, 0x00, 0xff, 0xd9]),
  ]);
}

function gif(width: number, height: number): Buffer {
  const size = Buffer.alloc(4);
  size.writeUInt16LE(width, 0);
  size.writeUInt16LE(height, 2);
  return Buffer.concat([Buffer.from("GIF89a"), size, Buffer.alloc(4)]);
}

const webp = (): Buffer => Buffer.concat([Buffer.from("RIFF"), Buffer.alloc(4), Buffer.from("WEBP"), Buffer.alloc(8)]);

describe("detectImage", () => {
  it("reads the type and the pixel size from the first bytes", () => {
    expect(detectImage(png(640, 480))).toEqual({ mimeType: "image/png", width: 640, height: 480 });
    expect(detectImage(jpeg(300, 200))).toEqual({ mimeType: "image/jpeg", width: 300, height: 200 });
    expect(detectImage(gif(16, 9))).toEqual({ mimeType: "image/gif", width: 16, height: 9 });
    expect(detectImage(webp())).toEqual({ mimeType: "image/webp" });
  });

  it("does not trust anything but the content", () => {
    expect(detectImage(Buffer.from("<svg xmlns='http://www.w3.org/2000/svg'/>"))).toBeNull();
    expect(detectImage(Buffer.from("not an image at all"))).toBeNull();
    expect(detectImage(Buffer.alloc(0))).toBeNull();
    expect(createAttachment("photo.png", Buffer.from("not an image"))).toMatchObject({ kind: "file" });
    expect(createAttachment("photo.png", png(2, 2))).toMatchObject({ kind: "image", mimeType: "image/png" });
    expect(createAttachment("photo.jpg", png(2, 2))).toMatchObject({ kind: "file" });
  });
});

describe("fitImage", () => {
  const natural = { width: 400, height: 200 };

  it("keeps the proportions when only one side is given", () => {
    expect(fitImage(natural, { width: 100 })).toEqual({ width: 100, height: 50 });
    expect(fitImage(natural, { height: 100 })).toEqual({ width: 200, height: 100 });
  });

  it("fits inside the box when both sides are given, without stretching", () => {
    expect(fitImage(natural, { width: 100, height: 100 })).toEqual({ width: 100, height: 50 });
    expect(fitImage(natural, { width: 1000, height: 100 })).toEqual({ width: 200, height: 100 });
  });

  it("keeps its own size, shrunk to a limit and never enlarged", () => {
    expect(fitImage(natural, {})).toEqual(natural);
    expect(fitImage(natural, {}, { width: 200, height: 1000 })).toEqual({ width: 200, height: 100 });
    expect(fitImage(natural, {}, { width: 4000, height: 4000 })).toEqual(natural);
  });

  it("refuses a picture without a size", () => {
    expect(() => fitImage({ width: 0, height: 0 }, {})).toThrow(/no size/);
  });

  it("reads pixels as points and centres a picture", () => {
    expect(naturalSize({ naturalWidth: 400, naturalHeight: 200 })).toEqual({ width: 300, height: 150 });
    expect(naturalSize({})).toBeNull();
    expect(centreIn({ left: 0, top: 0, width: 960, height: 540 }, { width: 300, height: 100 })).toEqual({ left: 330, top: 220, width: 300, height: 100 });
  });
});

describe("isPrivateAddress", () => {
  it.each(["127.0.0.1", "10.1.2.3", "172.16.0.1", "172.31.255.255", "192.168.1.1", "169.254.169.254", "100.64.0.1", "0.0.0.0", "224.0.0.1",
    "::1", "::", "fe80::1", "fc00::1", "fd12:3456::1", "ff02::1", "::ffff:127.0.0.1", "::ffff:7f00:1", "::ffff:10.0.0.1"])("%s is not public", (address) => {
    expect(isPrivateAddress(address)).toBe(true);
  });

  it.each(["8.8.8.8", "1.1.1.1", "172.32.0.1", "93.184.216.34", "2606:4700:4700::1111", "::ffff:8.8.8.8"])("%s is public", (address) => {
    expect(isPrivateAddress(address)).toBe(false);
  });
});

describe("downloadImage", () => {
  let server: http.Server;
  let base: string;
  const image = png(8, 4);

  beforeEach(async () => {
    server = http.createServer((request, response) => {
      if (request.url === "/ok.png") { response.writeHead(200, { "content-type": "image/png" }); response.end(image); }
      else if (request.url === "/redirect") { response.writeHead(302, { location: "/ok.png" }); response.end(); }
      else if (request.url === "/loop") { response.writeHead(302, { location: "/loop" }); response.end(); }
      else if (request.url === "/big") { response.writeHead(200, { "content-length": "5000000" }); response.end(Buffer.alloc(100)); }
      else if (request.url === "/stream") { response.writeHead(200); response.write(Buffer.alloc(3000)); response.end(Buffer.alloc(3000)); }
      else { response.writeHead(404); response.end("nope"); }
    });
    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  });
  afterEach(() => new Promise<void>((resolve) => { server.closeAllConnections(); server.close(() => resolve()); }));

  const open = { maxBytes: 1_000_000, allowPrivateNetwork: true };

  it("downloads a picture and follows a redirect", async () => {
    expect((await downloadImage(`${base}/ok.png`, open)).bytes.equals(image)).toBe(true);
    const redirected = await downloadImage(`${base}/redirect`, open);
    expect(redirected.bytes.equals(image)).toBe(true);
    expect(redirected.finalUrl).toBe(`${base}/ok.png`);
  });

  it("refuses what is too big, missing, or an endless redirect", async () => {
    await expect(downloadImage(`${base}/big`, { ...open, maxBytes: 1000 })).rejects.toThrow(/limit/);
    await expect(downloadImage(`${base}/stream`, { ...open, maxBytes: 4000 })).rejects.toThrow(/limit/);
    await expect(downloadImage(`${base}/missing`, open)).rejects.toThrow(/HTTP 404/);
    await expect(downloadImage(`${base}/loop`, open)).rejects.toThrow(/redirects/);
  });

  it("never reaches this computer or its network unless told to", async () => {
    const closed = { maxBytes: 1_000_000 };
    for (const url of [`${base}/ok.png`, base.replace("127.0.0.1", "localhost") + "/ok.png", `http://[::1]:${new URL(base).port}/ok.png`, "http://169.254.169.254/latest/meta-data/", "http://10.0.0.5/a.png"]) {
      await expect(downloadImage(url, closed), url).rejects.toBeInstanceOf(ImageDownloadError);
    }
  });

  it("takes http and https addresses only, without credentials", async () => {
    await expect(downloadImage("file:///C:/Windows/win.ini", open)).rejects.toThrow(/only http/);
    await expect(downloadImage("ftp://example.com/a.png", open)).rejects.toThrow(/only http/);
    await expect(downloadImage("not a url", open)).rejects.toThrow(/not a valid address/);
    await expect(downloadImage("https://user:secret@example.com/a.png", open)).rejects.toThrow(/user name or password/);
  });
});

describe("insert_image through the bridge", () => {
  let root: string;
  beforeEach(() => { root = fs.mkdtempSync(path.join(os.tmpdir(), "img-")); });
  afterEach(() => fs.rmSync(root, { recursive: true, force: true }));

  function setup(fetchImage = vi.fn(async () => ({ bytes: png(40, 20), finalUrl: "https://cdn.example.com/a/logo.png?size=big" }))) {
    const sessions = new WorkbookSessions({ ttlMs: 60_000 });
    const session = sessions.create("0123456789abcdef", null);
    const limits = { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 };
    const gateway = new OfficeRuntime(sessions, limits, 1000, undefined, false, null, new ReferenceFiles([root]), null, undefined, null, null, null, null, fetchImage);
    const sent: Array<{ tool: string; args: any }> = [];
    let approve = true;
    vi.spyOn(session, "requestTool").mockImplementation(async (tool: string, args: unknown) => {
      sent.push({ tool, args });
      return tool.startsWith("__authorize") ? { ok: true, data: { approved: approve } } : { ok: true, data: { summary: "placed" } };
    });
    const insert = (args: Record<string, unknown>) => gateway.invoke({ sessionId: session.id, tool: "insert_image", args });
    const forwarded = () => sent.filter((call) => call.tool === "insert_image").map((call) => call.args);
    return { session, gateway, insert, sent, forwarded, fetchImage, deny: () => { approve = false; } };
  }

  it("reads a data: address and hands the pane the picture and its size", async () => {
    const { insert, forwarded } = setup();
    const result = await insert({ source: `data:image/png;base64,${png(40, 20).toString("base64")}`, slide_id: "256", alt_text: "Logo" });
    expect(result).toMatchObject({ ok: true });
    expect(forwarded()).toHaveLength(1);
    expect(forwarded()[0]).toMatchObject({ mimeType: "image/png", naturalWidth: 40, naturalHeight: 20, sourceLabel: "pasted data", slide_id: "256", alt_text: "Logo" });
    expect(forwarded()[0].imageBase64).toBe(png(40, 20).toString("base64"));
  });

  it("reads plain base64, also wrapped over lines, and a JPEG", async () => {
    const { insert, forwarded } = setup();
    expect(await insert({ source: png(10, 10).toString("base64") })).toMatchObject({ ok: true });
    const wrapped = jpeg(30, 20).toString("base64").replace(/(.{20})/g, "$1\n");
    expect(await insert({ source: wrapped })).toMatchObject({ ok: true });
    expect(forwarded().map((args) => args.mimeType)).toEqual(["image/png", "image/jpeg"]);
    expect(forwarded()[1]).toMatchObject({ naturalWidth: 30, naturalHeight: 20 });
  });

  it("reads an attachment of the chat", async () => {
    const { session, insert, forwarded } = setup();
    const attachment = createAttachment("chart.png", png(5, 5));
    session.attachments.set(attachment.id, attachment);
    expect(await insert({ source: `attachment://${attachment.id}` })).toMatchObject({ ok: true });
    expect(forwarded()[0]).toMatchObject({ sourceLabel: "chart.png" });
    expect(await insert({ source: "attachment://00000000-0000-0000-0000-000000000000" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("reads a file in a folder the user opened up", async () => {
    const { insert, forwarded } = setup();
    const file = path.join(root, "photo.gif");
    fs.writeFileSync(file, gif(12, 6));
    expect(await insert({ source: file })).toMatchObject({ ok: true });
    expect(forwarded()[0]).toMatchObject({ mimeType: "image/gif", naturalWidth: 12, sourceLabel: "photo.gif" });
    expect(await insert({ source: path.join(root, "missing.png") })).toMatchObject({ ok: false });
  });

  it("asks before downloading, then sends only the picture on", async () => {
    const { insert, sent, forwarded, fetchImage } = setup();
    expect(await insert({ source: "https://cdn.example.com/a/logo.png?size=big" })).toMatchObject({ ok: true });
    expect(sent[0]).toEqual({ tool: "__authorize_image_url", args: { url: "https://cdn.example.com/a/logo.png?size=big" } });
    expect(fetchImage).toHaveBeenCalledWith("https://cdn.example.com/a/logo.png?size=big", expect.objectContaining({ maxBytes: 10 * 1024 * 1024 }));
    expect(forwarded()[0]).toMatchObject({ sourceLabel: "cdn.example.com/logo.png", naturalWidth: 40 });
  });

  it("downloads nothing when the user says no", async () => {
    const { insert, forwarded, fetchImage, deny } = setup();
    deny();
    expect(await insert({ source: "https://cdn.example.com/a.png" })).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(fetchImage).not.toHaveBeenCalled();
    expect(forwarded()).toHaveLength(0);
  });

  it("downloads nothing when the administrator turned the web off, but still reads other sources", async () => {
    const { session, insert, fetchImage, gateway } = setup();
    gateway.policy = { ...gateway.policy, allowWebSearch: false };
    expect(await insert({ source: "https://cdn.example.com/a.png" })).toMatchObject({ ok: false, error: { code: "unsupported", message: expect.stringContaining("administrator") } });
    expect(fetchImage).not.toHaveBeenCalled();
    expect(await insert({ source: png(4, 4).toString("base64") })).toMatchObject({ ok: true });
    void session;
  });

  it("explains what was wrong with a source that is not a usable picture", async () => {
    const { insert, forwarded } = setup(vi.fn(async () => ({ bytes: Buffer.from("<html>sign in</html>"), finalUrl: "https://example.com/login" })));
    const refused = async (source: string) => {
      const result = await insert({ source });
      expect(result.ok).toBe(false);
      return result.ok ? "" : result.error.message;
    };
    expect(await refused(`data:image/png;base64,${Buffer.from("hello world, no picture").toString("base64")}`)).toMatch(/not a PNG, JPEG or GIF/);
    expect(await refused(`data:image/webp;base64,${webp().toString("base64")}`)).toMatch(/WEBP.*convert it to PNG/);
    expect(await refused("data:image/png,hello")).toMatch(/base64 encoded/);
    expect(await refused("https://example.com/login")).toMatch(/not a PNG, JPEG or GIF/);
    // Just over 10 MB still fits the argument limit, so the bridge is the one to refuse it; far over never gets that far.
    expect(await refused(Buffer.concat([png(2, 2), Buffer.alloc(10 * 1024 * 1024 + 50 - png(2, 2).length)]).toString("base64"))).toMatch(/over the .* byte limit/);
    expect(await refused(Buffer.concat([png(2, 2), Buffer.alloc(11 * 1024 * 1024)]).toString("base64"))).toMatch(/invalid arguments/);
    expect(forwarded()).toHaveLength(0);
  });

  it("does not let the model supply the picture data itself", () => {
    expect(InsertImageInputSchema.safeParse({ source: "https://a.example/x.png", imageBase64: "AAAA" }).success).toBe(false);
    expect(InsertImageInputSchema.safeParse({ source: "https://a.example/x.png", width: 0 }).success).toBe(false);
    expect(InsertImageInputSchema.safeParse({ source: "https://a.example/x.png", anchor: "B2:D8", name: "logo" }).success).toBe(true);
  });
});

describe("insert_image in every host", () => {
  it("is offered to Excel, Word and PowerPoint, as a change to the file, with wording for each", () => {
    for (const host of HOST_IDS) expect(exposedToolNamesForHost(host, true), host).toContain("insert_image");
    expect(SHARED_WRITE_TOOLS).toContain("insert_image");
    const base = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "insert_image")!;
    expect(base.annotations).toMatchObject({ readOnlyHint: false });
    const word = definitionForHost("word", base).description;
    const powerpoint = definitionForHost("powerpoint", base).description;
    expect(word).toContain("paragraph");
    expect(word).not.toMatch(/workbook|cell/i);
    expect(powerpoint).toContain("slide");
    expect(definitionForHost("excel", base).description).toContain("sheet");
  });

  it("refuses a call that did not come through the bridge", async () => {
    expect(await executeSharedTool("excel", "insert_image", { source: "https://a.example/x.png" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });
});

describe("insert_image in Word", () => {
  let doc: FakeWordDoc;
  beforeEach(() => { doc = installWordFake(new FakeWordDoc(["Title", "First paragraph.", "Second paragraph."])); });
  afterEach(() => uninstallWordFake());

  const call = (extra: Partial<ResolvedInsertImageInput> = {}) => executeSharedTool("word", "insert_image", {
    source: "data:image/png;base64,AAAA", imageBase64: "AAAA", mimeType: "image/png", naturalWidth: 400, naturalHeight: 200, sourceLabel: "pasted data", ...extra,
  });

  it("puts the picture in a paragraph of its own at the end, in the normal style", async () => {
    doc.paragraphs[2]!.style = "Heading 1";
    doc.paragraphs[2]!.styleBuiltIn = "Heading1";
    const result = await call({ alt_text: "A logo" });
    expect(result).toMatchObject({ ok: true, data: { summary: expect.stringContaining("at the end"), paragraph_index: 3 } });
    expect(doc.paragraphs).toHaveLength(4);
    expect(doc.pictureLog).toHaveLength(1);
    expect(doc.pictureLog[0]).toMatchObject({ base64: "AAAA", location: "Start", altTextDescription: "A logo", lockAspectRatio: true });
    expect(doc.pictureLog[0]!.paragraph).toBe(doc.paragraphs[3]);
    expect(doc.paragraphs[3]!.styleBuiltIn).toBe("Normal");
  });

  it("goes after the paragraph asked for and takes the size asked for", async () => {
    const result = await call({ after_paragraph: 0, width: 100 });
    expect(result).toMatchObject({ ok: true, data: { width: 100, height: 50 } });
    expect(doc.paragraphs.map((p) => p.pictures)).toEqual([0, 1, 0, 0]);
    expect(doc.pictureLog[0]).toMatchObject({ width: 100, height: 50 });
  });

  it("uses the empty first paragraph of a new document instead of adding one", async () => {
    doc = installWordFake(new FakeWordDoc([""]));
    expect(await call()).toMatchObject({ ok: true, data: { summary: expect.stringContaining("empty document") } });
    expect(doc.paragraphs).toHaveLength(1);
    expect(doc.paragraphs[0]!.pictures).toBe(1);
  });

  it("says which paragraph is out of range", async () => {
    expect(await call({ after_paragraph: 9 })).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("after_paragraph 9") } });
    expect(doc.pictureLog).toHaveLength(0);
  });
});

describe("insert_image in Excel", () => {
  let book: MockWorkbook;
  let sheet: MockSheet;
  beforeEach(() => {
    book = new MockWorkbook();
    sheet = book.addSheet("Sheet1");
    book.addSheet("Other");
    book.selection = { sheet: "Sheet1", address: "A2:C2" };
    installOfficeMock(book);
  });

  const call = (extra: Partial<ResolvedInsertImageInput> = {}) => executeSharedTool("excel", "insert_image", {
    source: "data:image/png;base64,AAAA", imageBase64: "AAAA", mimeType: "image/png", naturalWidth: 400, naturalHeight: 200, sourceLabel: "pasted data", ...extra,
  });

  it("puts the picture at the selected cell, at its own size", async () => {
    const result = await call({ alt_text: "A logo" });
    expect(result).toMatchObject({ ok: true, data: { sheet: "Sheet1", left: 0, top: 15, width: 200, height: 100, source: "pasted data" } });
    const shape = sheet.shapeList[0]!;
    expect(shape.type).toBe("Image");
    expect(shape.name).toMatch(/^Image_/);
    expect(shape.altTextDescription).toBe("A logo");
  });

  it("fits it inside a cell range without stretching, centred", async () => {
    await call({ anchor: "B2:D4", name: "logo" });
    const shape = sheet.shapeList[0]!;
    expect([shape.name, shape.left, shape.top, shape.width, shape.height]).toEqual(["logo", 75, 15, 90, 45]);
  });

  it("takes explicit points, and keeps the proportions when only the width is given", async () => {
    await call({ left: 30, top: 40, width: 100 });
    const shape = sheet.shapeList[0]!;
    expect([shape.left, shape.top, shape.width, shape.height]).toEqual([30, 40, 100, 50]);
  });

  it("shrinks a big picture to the default limit but leaves a sized one alone", async () => {
    await call({ left: 0, top: 0, width: 900 });
    expect(sheet.shapeList[0]).toMatchObject({ width: 900, height: 450 });
  });

  it("goes to another sheet from its corner, and refuses a name that is taken", async () => {
    expect(await call({ sheet: { name: "Other" }, name: "pic" })).toMatchObject({ ok: true, data: { sheet: "Other", left: 0, top: 0 } });
    expect(await call({ sheet: { name: "Other" }, name: "pic" })).toMatchObject({ ok: false, error: { code: "invalid_arguments", message: expect.stringContaining("already exists") } });
  });
});

describe("insert_image in PowerPoint", () => {
  const fx = useFakeDeck();
  const call = (extra: Partial<ResolvedInsertImageInput> = {}) => executeSharedTool("powerpoint", "insert_image", {
    source: "data:image/png;base64,AAAA", imageBase64: "AAAA", mimeType: "image/png", naturalWidth: 400, naturalHeight: 200, sourceLabel: "pasted data", slide_id: fx().agenda, ...extra,
  });

  it("shows the slide, puts the picture in the middle of it at its own size and names it", async () => {
    const result = await call({ name: "Sales chart", alt_text: "Sales by region" });
    const picture = fx().deck.slideById(fx().agenda).shapes.at(-1)!;
    expect(result).toMatchObject({ ok: true, data: { slide_id: fx().agenda, ref: `${fx().agenda}:${picture.id}`, name: "Sales chart", left: 330, top: 195, width: 300, height: 150 } });
    expect(fx().deck.pictureRequests).toEqual([{ slideId: fx().agenda, data: "AAAA", options: { coercionType: "image", imageLeft: 330, imageTop: 195, imageWidth: 300, imageHeight: 150 } }]);
    expect(picture).toMatchObject({ name: "Sales chart", type: "Image", altDescr: "Sales by region" });
    expect(fx().deck.selection.slideIds).toEqual([fx().agenda]);
  });

  it("takes the place and size asked for, keeping the proportions", async () => {
    await call({ left: 10, top: 20, width: 150 });
    expect(fx().deck.pictureRequests[0]!.options).toMatchObject({ imageLeft: 10, imageTop: 20, imageWidth: 150, imageHeight: 75 });
  });

  it("fits a picture bigger than the slide inside it", async () => {
    await call({ naturalWidth: 4000, naturalHeight: 2000 });
    expect(fx().deck.pictureRequests[0]!.options).toMatchObject({ imageLeft: 0, imageTop: 30, imageWidth: 960, imageHeight: 480 });
  });

  it("uses the selected slide when none is named", async () => {
    fx().deck.selection = { slideIds: [fx().data], shapes: [] };
    expect(await call({ slide_id: undefined })).toMatchObject({ ok: true, data: { slide_id: fx().data } });
  });

  it("stops with a clear message when PowerPoint cannot take a picture, or refuses it", async () => {
    fx().deck.imageCoercion = false;
    expect(await call()).toMatchObject({ ok: false, error: { code: "unsupported", message: expect.stringContaining("ImageCoercion") } });
    fx().deck.imageCoercion = true;
    fx().deck.failPicture = true;
    expect(await call()).toMatchObject({ ok: false, error: { code: "office_error", message: expect.stringContaining("did not take the picture") } });
    expect(await call({ slide_id: "999" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
  });

  it("will not guess a size for a picture that does not say how big it is", async () => {
    expect(await call({ naturalWidth: undefined, naturalHeight: undefined, width: 100 })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(fx().deck.pictureRequests).toHaveLength(0);
  });
});
