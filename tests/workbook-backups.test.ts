import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { BackupError, MAX_BACKUP_BYTES, WorkbookBackups, safeBaseName } from "../bridge/src/workbook-backups.js";
import { PaneClient, getFreePort } from "./helpers.js";

const zip = (extra = 8) => Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.alloc(extra, 1)]);

let folder = "";
beforeEach(() => { folder = fs.mkdtempSync(path.join(os.tmpdir(), "backups-")); });
afterEach(() => fs.rmSync(folder, { recursive: true, force: true }));

describe("WorkbookBackups", () => {
  it("saves a workbook copy as an .xlsx named after the workbook and the time", () => {
    const store = new WorkbookBackups(path.join(folder, "b"), 10, () => new Date("2026-11-04T15:02:37.123Z"));
    const saved = store.save("Q3 Budget.xlsx", zip());
    expect(saved.file).toBe("Q3 Budget-20261104-150237.xlsx");
    expect(fs.readFileSync(path.join(folder, "b", saved.file)).equals(zip())).toBe(true);
    expect(store.list().map((b) => b.file)).toEqual([saved.file]);
  });

  it("never overwrites an earlier copy taken in the same second", () => {
    const store = new WorkbookBackups(folder, 10, () => new Date("2026-11-04T15:02:37Z"));
    const first = store.save("Book1", zip(1));
    const second = store.save("Book1", zip(2));
    expect(second.file).not.toBe(first.file);
    expect(fs.readdirSync(folder).sort()).toEqual([first.file, second.file].sort());
    expect(fs.readFileSync(path.join(folder, first.file)).length).toBe(5);
  });

  it("keeps only the newest copies", () => {
    let minute = 0;
    const store = new WorkbookBackups(folder, 3, () => new Date(Date.UTC(2026, 10, 4, 15, minute++, 0)));
    const files = Array.from({ length: 5 }, () => store.save("Book1", zip()).file);
    expect(store.list().map((b) => b.file).sort()).toEqual(files.slice(2).sort());
    expect(fs.readdirSync(folder).sort()).toEqual(files.slice(2).sort());
  });

  it("refuses data that is not a workbook, an empty copy, or one over the limit", () => {
    const store = new WorkbookBackups(folder);
    expect(() => store.save("Book1", Buffer.from("hello"))).toThrow(BackupError);
    expect(() => store.save("Book1", Buffer.alloc(0))).toThrow(/empty/);
    expect(() => store.save("Book1", Buffer.concat([zip(), Buffer.alloc(MAX_BACKUP_BYTES)]))).toThrow(/limit/);
    expect(fs.readdirSync(folder)).toEqual([]);
  });

  it("turns a workbook name into a safe file name", () => {
    expect(safeBaseName("..\\..\\evil:name?.xlsx")).toBe("_.._evil_name_");
    expect(safeBaseName("")).toBe("workbook");
    expect(safeBaseName("   ")).toBe("workbook");
    expect(safeBaseName("báo cáo tháng 10.xlsm")).toBe("báo cáo tháng 10");
    expect(safeBaseName("x".repeat(200)).length).toBe(60);
    expect(safeBaseName(".hidden")).toBe("hidden");
  });

  it("lists nothing before the first backup", () => {
    expect(new WorkbookBackups(path.join(folder, "missing")).list()).toEqual([]);
  });
});

describe("POST /session/backup", () => {
  let bridge: Bridge | null = null;
  let pane: PaneClient | null = null;
  afterEach(async () => { await pane?.close(); pane = null; await bridge?.close(); bridge = null; });

  async function connect() {
    const port = await getFreePort();
    bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error", configFile: path.join(folder, "config.json") }), {
      usageFile: null, savedReferencesFile: null, workspaceConfigPath: null, skillUploadDir: path.join(folder, "skills"),
      systemPromptFile: path.join(folder, "system-prompt.txt"), registry: new ProviderRegistry().register(new MockAdapter({ id: "claude" })),
    });
    const { url } = await bridge.listen();
    pane = await PaneClient.connect(url.replace("http", "ws") + "/ws");
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "backup-workbook" });
    const hello = await pane.waitFor((message) => message.type === "hello_ok");
    if (hello.type !== "hello_ok") throw new Error("no handshake");
    return { url, token: hello.sessionToken };
  }

  it("stores the copy next to the other Excel Partner data and says where", async () => {
    const { url, token } = await connect();
    const response = await fetch(url + "/session/backup", { method: "POST", headers: { authorization: `Bearer ${token}`, "x-file-name": encodeURIComponent("Báo cáo.xlsx") }, body: zip(64) });
    const body = await response.json() as { file: string; size: number; folder: string };
    expect(response.status).toBe(200);
    expect(body.folder).toBe(path.join(folder, "backups"));
    expect(body.file).toMatch(/^Báo cáo-\d{8}-\d{6}\.xlsx$/);
    expect(body.size).toBe(68);
    expect(fs.readFileSync(path.join(body.folder, body.file)).equals(zip(64))).toBe(true);
  });

  it("rejects callers without a live session and files that are not workbooks", async () => {
    const { url, token } = await connect();
    expect((await fetch(url + "/session/backup", { method: "POST", headers: { authorization: "Bearer nope", "x-file-name": "a.xlsx" }, body: zip() })).status).toBe(401);
    const notExcel = await fetch(url + "/session/backup", { method: "POST", headers: { authorization: `Bearer ${token}`, "x-file-name": "a.xlsx" }, body: "plain text" });
    expect(notExcel.status).toBe(400);
    expect(((await notExcel.json()) as { error: string }).error).toMatch(/not a workbook/);
    expect(fs.existsSync(path.join(folder, "backups"))).toBe(false);
  });
});
