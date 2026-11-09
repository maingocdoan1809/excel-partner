import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { BackupError, DocumentBackups, backupExtensionFor, safeBaseName } from "../bridge/src/document-backups.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

const zip = (extra = 8) => Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.alloc(extra, 1)]);
let folder = "";
beforeEach(() => { folder = fs.mkdtempSync(path.join(os.tmpdir(), "doc-backups-")); });
afterEach(() => fs.rmSync(folder, { recursive: true, force: true }));

describe("DocumentBackups", () => {
  it("keeps the file type of the host and lists all three kinds together", () => {
    let minute = 0;
    const store = new DocumentBackups(folder, 10, () => new Date(Date.UTC(2026, 10, 4, 15, minute++, 0)));
    const sheet = store.save("Budget.xlsx", zip());
    const doc = store.save("Plan Q3.docx", zip(), ".docx");
    const deck = store.save("Pitch.pptx", zip(), ".pptx");
    expect(sheet.file).toMatch(/^Budget-\d{8}-\d{6}\.xlsx$/);
    expect(doc.file).toMatch(/^Plan Q3-\d{8}-\d{6}\.docx$/);
    expect(deck.file).toMatch(/^Pitch-\d{8}-\d{6}\.pptx$/);
    expect(store.list().map((b) => b.file).sort()).toEqual([sheet.file, doc.file, deck.file].sort());
  });

  it("prunes the oldest across file types", () => {
    let minute = 0;
    const store = new DocumentBackups(folder, 2, () => new Date(Date.UTC(2026, 10, 4, 15, minute++, 0)));
    store.save("a", zip(), ".docx");
    const b = store.save("b", zip(), ".pptx");
    const c = store.save("c", zip(), ".xlsx");
    expect(store.list().map((x) => x.file).sort()).toEqual([b.file, c.file].sort());
  });

  it("names refusals after the kind of file and falls back to a matching base name", () => {
    const store = new DocumentBackups(folder);
    expect(() => store.save("x", Buffer.from("text"), ".docx")).toThrow(/not a document file/);
    expect(() => store.save("x", Buffer.from("text"), ".pptx")).toThrow(/not a presentation file/);
    expect(() => store.save("x", Buffer.alloc(0), ".docx")).toThrow(/document copy is empty/);
    expect(() => store.save("x", Buffer.from("text"), ".docx")).toThrow(BackupError);
    expect(store.save("", zip(), ".docx").file).toMatch(/^document-/);
    expect(store.save("", zip(), ".pptx").file).toMatch(/^presentation-/);
    expect(store.save("", zip()).file).toMatch(/^workbook-/);
  });

  it("strips Word and PowerPoint extensions from the name", () => {
    expect(safeBaseName("Report.docx", "document")).toBe("Report");
    expect(safeBaseName("Deck.PPTX", "presentation")).toBe("Deck");
    expect(safeBaseName("Macros.pptm")).toBe("Macros");
    expect(backupExtensionFor("word")).toBe(".docx");
    expect(backupExtensionFor("powerpoint")).toBe(".pptx");
    expect(backupExtensionFor("excel")).toBe(".xlsx");
  });
});

describe("POST /session/backup for each host", () => {
  let bridge: Bridge | null = null;
  let pane: PaneClient | null = null;
  afterEach(async () => { await pane?.close(); pane = null; await bridge?.close(); bridge = null; });

  async function connect(host: string) {
    const port = await getFreePort();
    bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error", configFile: path.join(folder, "config.json") }), {
      usageFile: null, savedReferencesFile: null, workspaceConfigPath: null, skillUploadDir: path.join(folder, "skills"),
      systemPromptFile: path.join(folder, "system-prompt.txt"), registry: new ProviderRegistry().register(new MockAdapter({ id: "claude" })),
    });
    const { url } = await bridge.listen();
    pane = await PaneClient.connect(url.replace("http", "ws") + "/ws");
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: `backup-${host}-0001`, clientInfo: { host } as never });
    const hello = await pane.waitFor((message) => message.type === "hello_ok");
    if (hello.type !== "hello_ok") throw new Error("no handshake");
    return { url, token: hello.sessionToken };
  }

  it.each([["excel", "xlsx"], ["word", "docx"], ["powerpoint", "pptx"]])("saves a %s copy as .%s", async (host, extension) => {
    const { url, token } = await connect(host);
    const response = await fetch(url + "/session/backup", { method: "POST", headers: { authorization: `Bearer ${token}`, "x-file-name": encodeURIComponent("Báo cáo") }, body: zip(32) });
    const body = await response.json() as { file: string; folder: string };
    expect(response.status).toBe(200);
    expect(body.file).toMatch(new RegExp(`^Báo cáo-\\d{8}-\\d{6}\\.${extension}$`));
    expect(fs.readFileSync(path.join(body.folder, body.file)).equals(zip(32))).toBe(true);
  });

  it("refuses data that is not a document in the host's own words", async () => {
    const { url, token } = await connect("word");
    const response = await fetch(url + "/session/backup", { method: "POST", headers: { authorization: `Bearer ${token}`, "x-file-name": "a.docx" }, body: "plain text" });
    expect(response.status).toBe(400);
    expect(((await response.json()) as { error: string }).error).toMatch(/not a document file/);
  });
});
