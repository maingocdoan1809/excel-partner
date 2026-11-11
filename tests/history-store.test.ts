import Database from "better-sqlite3";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { HistoryStore } from "../bridge/src/history-store.js";
import type { StoredSession } from "../protocol/src/history.js";
const dirs: string[] = [], stores: HistoryStore[] = [];
afterEach(() => { for (const store of stores.splice(0)) store.close(); for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function fixture() { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "history-db-")); dirs.push(dir); const file = path.join(dir, "history.db"), store = new HistoryStore(file); stores.push(store); return { dir, file, store }; }
const chat = (id = "chat"): StoredSession => ({ id, host: "excel", provider: "copilot", providerSessionId: "native-1", title: "Question", startedAt: 1, updatedAt: 2,
  messages: [{ kind: "user", text: "hello" }, { kind: "tool_group", text: "", tools: [{ name: "read_range", args: "{}", result: "answer", state: "ok", error: null }] }, { kind: "artifact", text: "", artifact: { name: "report.html", content: "<p>report</p>" } }], changes: [{ at: 2, tool: "write_range", ok: true, summary: "Written", target: null }] });
const input = (session = chat(), baseRevision = 0, operationId = "operation-1", generation = 0) => ({ session, baseRevision, operationId, generation });
describe("durable history acceptance", () => {
  it("H11 rolls back earlier imported records if SQLite rejects a later record", () => {
    const w = fixture(); const connection = new Database(w.file);
    connection.exec("CREATE TRIGGER fail_import BEFORE INSERT ON history_chats WHEN NEW.id='bad' BEGIN SELECT RAISE(ABORT,'import failed'); END"); connection.close();
    expect(() => w.store.importLegacy([chat("first"), chat("bad")])).toThrow("import failed");
    expect(w.store.list("excel").total).toBe(0);
  });

  it("H05 enforces revisions across independent SQLite connections", () => {
    const w = fixture(); w.store.save("excel", input()); const second = new HistoryStore(w.file); stores.push(second);
    const before = w.store.get("excel", "chat")!;
    const updated = chat(); updated.messages.push({ kind: "assistant", text: "other connection" });
    second.save("excel", input(updated, before.revision, "other"));
    expect(() => w.store.save("excel", input(before.session, before.revision, "stale"))).toThrow("HISTORY_CONFLICT");
  });

  it("H18 reports SQLite write failures while preserving the previously committed record", () => {
    const w = fixture(); w.store.save("excel", input());
    const connection = new Database(w.file);
    connection.exec("CREATE TRIGGER fail_write BEFORE UPDATE ON history_chats BEGIN SELECT RAISE(ABORT,'simulated storage failure'); END");
    connection.close();
    const changed = chat(); changed.messages.push({ kind: "assistant", text: "unsaved" });
    expect(() => w.store.save("excel", input(changed, 1, "failed"))).toThrow("simulated storage failure");
    expect(w.store.get("excel", "chat")?.session).toEqual(chat());
  });
  it("H19 refuses a future schema without changing its version or data", () => {
    const w = fixture(); w.store.save("excel", input()); w.store.close();
    const connection = new Database(w.file); connection.pragma("user_version = 99"); connection.close();
    expect(() => new HistoryStore(w.file)).toThrow("HISTORY_SCHEMA_NEWER");
    const reader = new Database(w.file); expect(reader.pragma("user_version", { simple: true })).toBe(99); expect(reader.prepare("SELECT count(*) n FROM history_chats").get()).toMatchObject({ n: 1 }); reader.close();
  });

  it("H01/H02 preserves full data and native identity across reopening SQLite", () => {
    const w = fixture(), session = chat();
    expect(w.store.save("excel", input(session))).toMatchObject({ revision: 1 });
    w.store.close(); const reopened = new HistoryStore(w.file); stores.push(reopened);
    expect(reopened.get("excel", session.id)?.session).toEqual(session);
    expect(fs.readFileSync(w.file).subarray(0, 15).toString()).toBe("SQLite format 3");
  });
  it("H03/H04 retains large history without pruning and pages summaries/transcripts", () => {
    const { store } = fixture();
    for (let i = 0; i < 45; i++) store.save("excel", input(chat(String(i)), 0, `create-${i}`));
    const long = chat("long"); long.messages = Array.from({ length: 405 }, (_, i) => ({ kind: "user", text: i === 0 ? "x".repeat(25_000) : String(i) + "x".repeat(8_000) }));
    store.save("excel", input(long));
    expect(store.list("excel", 0, 10)).toMatchObject({ total: 46, items: expect.any(Array) });
    expect(store.list("excel", 0, 10).items).toHaveLength(10);
    expect(store.page("excel", "long", 0, 100)?.session.messages[0]?.text).toHaveLength(25_000);
    expect(store.page("excel", "long", 400, 100)?.session.messages).toHaveLength(5);
  });
  it("H05/H06 rejects stale snapshots but acknowledges identical retries", () => {
    const { store } = fixture(), first = input();
    expect(store.save("excel", first).revision).toBe(1);
    expect(store.save("excel", first).revision).toBe(1);
    const newer = chat(); newer.messages.push({ kind: "assistant", text: "new" });
    store.save("excel", input(newer, 1, "operation-2"));
    expect(() => store.save("excel", input(chat(), 1, "stale"))).toThrow("HISTORY_CONFLICT");
    expect(store.get("excel", "chat")?.session.messages.at(-1)?.text).toBe("new");
    expect(() => store.save("excel", input(newer, 1, "operation-2", 0))).not.toThrow();
  });
  it("H07 tombstones deleted chats and advances generation on clear", () => {
    const { store } = fixture(); store.save("excel", input()); store.delete("excel", "chat");
    expect(() => store.save("excel", input())).toThrow("HISTORY_DELETED");
    expect(store.get("excel", "chat")).toBeNull();
    store.clear("excel");
    expect(() => store.save("excel", input(chat("offline-new")))).toThrow("HISTORY_GENERATION");
    expect(store.save("excel", input(chat("fresh"), 0, "fresh", 1)).revision).toBe(1);
  });
  it("H08 isolates hosts even when a client forges the payload host", () => {
    const { store } = fixture(); store.save("word", input(chat()));
    expect(store.get("excel", "chat")).toBeNull();
    expect(store.get("word", "chat")?.session.host).toBe("word");
    store.clear("excel"); expect(store.list("word").total).toBe(1);
  });
  it("H10 imports all hosts idempotently without replacing a newer chat or tombstone", () => {
    const { store } = fixture();
    store.importLegacy([chat(), { ...chat("word"), host: "word" }]);
    const newer = { ...chat("word"), host: "word" as const }; newer.messages.push({ kind: "assistant", text: "newer" });
    store.save("word", input(newer, 1, "newer"));
    store.delete("excel", "chat"); store.importLegacy([chat(), { ...chat("word"), host: "word" }]);
    expect(store.list("excel").total).toBe(0); expect(store.list("word").total).toBe(1); expect(store.get("word", "word")?.session.messages.at(-1)?.text).toBe("newer");
  });
  it("H11 rolls back the whole import if a later record is invalid", () => {
    const { store } = fixture();
    expect(() => store.importLegacy([chat(), { id: "bad" } as StoredSession])).toThrow();
    expect(store.list("excel").total).toBe(0);
  });
  it("H19 preserves corrupt database bytes and refuses unknown future schemas", () => {
    const w = fixture(); w.store.close(); fs.writeFileSync(w.file, "broken database");
    expect(() => new HistoryStore(w.file)).toThrow(); expect(fs.readFileSync(w.file, "utf8")).toBe("broken database");
  });
  it("H20/H21 exports, and only deletes explicit old chats", () => {
    const w = fixture(); w.store.save("excel", input());
    expect(w.store.export("excel")).toEqual([chat()]);
    expect(w.store.cleanup("excel", 3, ["chat"])).toBe(0);
    expect(w.store.cleanup("excel", 3, [])).toBe(1);
    expect(w.store.stats("excel")).toMatchObject({ count: 0, generation: 0 });
  });
  it("reports an unusable SQLite driver as HISTORY_DRIVER_UNAVAILABLE and passes other failures through", () => {
    const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "history-driver-")), "history.db");
    const missing = () => { throw Object.assign(new Error("Could not locate the bindings file. Tried: better_sqlite3.node"), { code: "ERR_DLOPEN_FAILED" }); };
    expect(() => new HistoryStore(file, missing)).toThrowError(expect.objectContaining({ code: "HISTORY_DRIVER_UNAVAILABLE" }));
    const other = () => { throw new Error("EACCES: permission denied"); };
    expect(() => new HistoryStore(file, other)).toThrow("EACCES");
  });
});
