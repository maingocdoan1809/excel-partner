import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { PassThrough } from "node:stream";
import type http from "node:http";
import { afterEach, describe, expect, it } from "vitest";
import { HistoryRoutes } from "../bridge/src/routes/history.js";
import { HistoryStore, maintainHistory } from "../bridge/src/history-store.js";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { Logger } from "../bridge/src/logger.js";

let dir = "";
afterEach(() => { if (dir) fs.rmSync(dir, { recursive: true, force: true }); });

function fixture(openStore?: (file: string) => HistoryStore) {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), "history-restart-"));
  const lines: string[] = [];
  const sessions = new WorkbookSessions({ ttlMs: 60_000 });
  const session = sessions.create("0123456789abcdef", { host: "excel" } as never);
  session.attachSocket({ readyState: 1, send() {}, close() {} } as never);
  const routes = new HistoryRoutes(sessions, path.join(dir, "history.db"), new Logger("history", "debug", line => lines.push(line)), openStore);
  const request = (method: string, query = "") => {
    const req = Object.assign(new PassThrough(), { method, headers: { authorization: `Bearer ${session.sessionToken}` } }) as unknown as http.IncomingMessage & PassThrough;
    let status = 0, body = "";
    const res = { headersSent: false, writeHead(code: number) { status = code; }, end(text: string) { body = text; } } as unknown as http.ServerResponse;
    const done = routes.handle(req, res, new URL(`https://localhost/session/history${query}`)).then(() => ({ status, body: JSON.parse(body) }));
    return { req, done };
  };
  return { routes, request, lines };
}

const save = { action: "save", baseRevision: 0, generation: 0, operationId: "op1",
  session: { id: "chat", provider: "copilot", title: "t", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "hi" }], changes: [] } };

describe("history during a release switch", () => {
  it("answers a request that was in flight when the store closed with HISTORY_RESTARTING, without an error log", async () => {
    const w = fixture();
    const list = w.request("GET"); list.req.end();
    expect((await list.done).status).toBe(200);
    const post = w.request("POST");
    const text = JSON.stringify(save);
    post.req.write(text.slice(0, 10));
    w.routes.close();
    post.req.end(text.slice(10));
    expect(await post.done).toEqual({ status: 503, body: expect.objectContaining({ code: "HISTORY_RESTARTING" }) });
    const later = w.request("GET"); later.req.end();
    expect((await later.done).body.code).toBe("HISTORY_RESTARTING");
    expect(w.lines.filter(line => / (ERROR|WARN) /.test(line))).toEqual([]);
  });

  it("a closed store throws a typed error instead of a driver error", () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "history-restart-"));
    const store = new HistoryStore(path.join(dir, "history.db"));
    store.close();
    expect(() => store.list("excel")).toThrow(expect.objectContaining({ code: "HISTORY_CLOSED" }));
  });

  it("reports HISTORY_DRIVER_UNAVAILABLE when the SQLite module cannot be loaded at all", async () => {
    const missing = () => { throw Object.assign(new Error("Cannot find module 'better-sqlite3'\nRequire stack:\n- x"), { code: "MODULE_NOT_FOUND" }); };
    const w = fixture(file => new HistoryStore(file, missing));
    const list = w.request("GET"); list.req.end();
    expect(await list.done).toEqual({ status: 503, body: expect.objectContaining({ code: "HISTORY_DRIVER_UNAVAILABLE" }) });
  });
});

describe("history housekeeping", () => {
  // Backup age is read from the file's real modification time, so the test runs at the real time.
  const DAY = 24 * 60 * 60_000, NOW = Date.now();
  const chat = (id: string, updatedAt: number) => ({ session: { id, provider: "copilot", title: id, startedAt: updatedAt, updatedAt, messages: [{ kind: "user", text: "hi" }], changes: [] } as never,
    baseRevision: 0, generation: 0, operationId: id });

  it("applies the retention policy and keeps two daily backups", async () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "history-maintain-"));
    const file = path.join(dir, "history.db");
    expect(await maintainHistory(file, 30, NOW)).toEqual({ removed: 0, backedUp: false });
    expect(fs.existsSync(file)).toBe(false);
    const store = new HistoryStore(file);
    store.save("excel", chat("old", NOW - 31 * DAY)); store.save("word", chat("older", NOW - 90 * DAY)); store.save("excel", chat("new", NOW - DAY));
    store.close();
    expect(await maintainHistory(file, 30, NOW)).toEqual({ removed: 2, backedUp: true });
    const backup = new HistoryStore(file + ".bak");
    expect(backup.list("excel").items.map(item => item.id)).toEqual(["new"]);
    expect(backup.stats("word").count).toBe(0);
    backup.close();
    expect(await maintainHistory(file, undefined, NOW + 1000)).toEqual({ removed: 0, backedUp: false });
    fs.utimesSync(file + ".bak", new Date(NOW - 2 * DAY), new Date(NOW - 2 * DAY));
    expect(await maintainHistory(file, 0, NOW)).toEqual({ removed: 0, backedUp: true });
    expect(fs.existsSync(file + ".bak.1")).toBe(true);
  });
});
