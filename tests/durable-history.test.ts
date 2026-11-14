// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { DurableHistory } from "../frontend/taskpane/ui/durable-history.js";
import { MAX_STORED_TEXT } from "../frontend/taskpane/ui/sessions.js";
import type { StoredSession, HistorySave } from "../protocol/src/history.js";
const session = (): StoredSession => ({ id: "chat", title: "hello", provider: "copilot", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "hello" }], changes: [] });
afterEach(() => { localStorage.clear(); vi.useRealTimers(); vi.restoreAllMocks(); });
function fixture() {
 const operations = new Map<string, { revision: number; generation: number }>();
 const rows = new Map<string, { session: StoredSession; revision: number }>();
 const api = { historyStats: vi.fn(async () => ({ count: rows.size, bytes: 100, generation: 0 })), historyList: vi.fn(async () => ({ items: [...rows.values()].map(r => ({ ...r.session, host: "excel" as const, messageCount: r.session.messages.length, changeCount: 0, revision: r.revision })), total: rows.size, generation: 0 })),
 historyImport: vi.fn(async (_sessions: StoredSession[]) => ({ imported: 0 })), historyPage: vi.fn(async (id: string) => { const r = rows.get(id); if (!r) throw new Error("HISTORY_NOT_FOUND"); return { ...r, generation: 0, totalMessages: r.session.messages.length, totalChanges: 0 }; }),
 historySave: vi.fn(async (s: HistorySave) => { if (operations.has(s.operationId)) return operations.get(s.operationId)!; const old = rows.get(s.session.id); if ((old?.revision ?? 0) !== s.baseRevision) throw new Error("HISTORY_CONFLICT"); const revision = (old?.revision ?? 0) + 1; rows.set(s.session.id, { session: s.session, revision }); const ack = { revision, generation: 0 }; operations.set(s.operationId, ack); return ack; }),
 historyDelete: vi.fn(async (id?: string) => { if (id) rows.delete(id); else rows.clear(); return { count: rows.size, bytes: 100, generation: 0 }; }) };
 const notice = vi.fn(); const store = new DurableHistory(() => api, () => "excel", localStorage, notice);
 return { api, rows, notice, store };
}
describe("durable pane outbox", () => {
 it("H35 retries a temporary save failure automatically without reconnecting or opening Settings", async () => {
   vi.useFakeTimers(); const w = fixture(); await w.store.connect();
   w.api.historySave.mockRejectedValueOnce(new Error("temporary offline")); w.store.save(session());
   await expect(w.store.flush()).rejects.toThrow("temporary offline");
   await vi.advanceTimersByTimeAsync(1000); expect(w.store.pendingCount).toBe(0); expect(w.rows.has("chat")).toBe(true);
 });
 it("H29 continues saving an imported legacy chat without asking the user to resolve its initial revision", async () => {
   const old = session(); localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([old]));
   const w = fixture(); w.rows.set("chat", { session: old, revision: 1 }); await w.store.connect();
   const continued = session(); continued.messages.push({ kind: "assistant", text: "continued" });
   w.store.save(continued); await w.store.flush();
   expect(w.rows.get("chat")?.session.messages).toHaveLength(2); expect(w.store.pendingCount).toBe(0);
 });
 it("H30 ignores an old acknowledged outbox snapshot after another pane saved a longer transcript", async () => {
   const w = fixture(); const newer = session(); newer.messages.push({ kind: "assistant", text: "newer answer" });
   w.rows.set("chat", { session: newer, revision: 3 });
   localStorage.setItem("excelPartner.historyOutbox.v1.closed-pane", JSON.stringify([{ session: session(), host: "excel", baseRevision: 0, generation: 0, operationId: "old-op" }]));
   await w.store.connect(); const reloaded = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice);
   await reloaded.connect(); expect(reloaded.list()).toHaveLength(1); expect(reloaded.pendingCount).toBe(0);
   expect(w.rows.get("chat")?.session.messages).toHaveLength(2);
 });
 it("H31 keeps history readable when retrying an unrelated failed pending save", async () => {
   const w = fixture(); await w.store.connect(); w.store.save(session()); w.api.historySave.mockRejectedValue(new Error("disk unavailable"));
   await expect(w.store.connect()).resolves.toBeUndefined(); expect(w.store.pendingCount).toBe(1);
   await w.store.refresh(); expect(w.notice.mock.calls.some(([s]) => String(s).includes("History is not fully saved or available"))).toBe(false);
 });
 it("H27 can list acknowledged history after an initial connection failure", async () => {
   const w = fixture(); w.rows.set("chat", { session: session(), revision: 1 });
   w.api.historyStats.mockRejectedValueOnce(new Error("offline")); await expect(w.store.connect()).rejects.toThrow("offline");
   await w.store.refresh(); expect(w.store.list().map(s => s.id)).toEqual(["chat"]);
 });
 it("H17 replaces a provisional history item with a single native conversation", async () => {
   const w = fixture(); await w.store.connect(); w.store.save(session());
   w.store.save({ ...session(), messages: [] });
   w.store.save({ ...session(), id: "copilot:native", providerSessionId: "native" });
   await w.store.flush(); expect([...w.rows.keys()]).toEqual(["copilot:native"]);
 });

 it("H12 retains corrupt legacy bytes while allowing new SQLite-backed chats", async () => {
   localStorage.setItem("excelPartner.sessions.v1", "{broken"); const w = fixture();
   await w.store.connect(); w.store.save(session()); await w.store.flush();
   expect(w.rows.has("chat")).toBe(true); expect(localStorage.getItem("excelPartner.sessions.v1")).toBe("{broken");
   expect(localStorage.getItem("excelPartner.historyMigrated.v1.excel")).toBeNull();
 });

 it("H12 saves to SQLite when browser storage cannot even be read", async () => {
   const w = fixture(); const blocked = { getItem() { throw new Error("blocked"); }, setItem() { throw new Error("blocked"); }, removeItem() {} };
   const store = new DurableHistory(() => w.api, () => "excel", blocked, w.notice);
   await store.connect(); store.save(session()); await store.flush(); expect(w.rows.has("chat")).toBe(true);
 });
 it("H08 keeps identical legacy IDs isolated across Office hosts", () => {
   localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([{ ...session(), host: "excel", title: "Excel" }, { ...session(), host: "word", title: "Word" }]));
   const w = fixture(); expect(w.store.get("chat")?.title).toBe("Excel");
 });

 it("H06 retries the exact operation after its successful HTTP acknowledgement was lost", async () => {
   const w = fixture(); await w.store.connect(); const save = w.api.historySave.getMockImplementation()!;
   w.api.historySave.mockImplementationOnce(async input => { await save(input); throw new Error("response lost"); });
   w.store.save(session()); await expect(w.store.flush()).rejects.toThrow("response lost");
   const reloaded = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice);
   await reloaded.connect(); expect(w.rows.get("chat")?.revision).toBe(1); expect(reloaded.pendingCount).toBe(0);
 });
 it("H32 saves divergent versions separately and reuses the saved copy for subsequent updates", async () => {
   const w = fixture(); await w.store.connect(); const original = session(); original.messages[0]!.text = "newer server version";
   w.rows.set("chat", { session: original, revision: 5 }); w.store.save(session()); w.store.save({ ...session(), id: "unrelated" });
   await w.store.flush(); expect(w.rows.has("unrelated")).toBe(true);
   const recovered = [...w.rows.values()].find(r => r.session.id.startsWith("recovered-"))!;
   expect(recovered.session.providerSessionId).toBeUndefined();
   expect(w.rows.get("chat")?.session.messages[0]?.text).toBe("newer server version");
   expect(recovered.session.messages[0]?.text).toBe("hello"); expect(w.store.pendingCount).toBe(0);
   const continued = session(); continued.messages.push({ kind: "assistant", text: "continued local copy" }); w.store.save(continued); await w.store.flush();
   expect(w.rows.size).toBe(3); expect(w.rows.get(recovered.session.id)?.session.messages).toHaveLength(2);
   const reloaded = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice); await reloaded.connect();
   expect((await reloaded.loadActive("chat"))?.id).toBe(recovered.session.id);
 });

 it("H05 summary refresh cannot rebase an older loaded transcript over another window's edits", async () => {
   const w = fixture(); w.rows.set("chat", { session: session(), revision: 1 }); await w.store.connect();
   const old = await w.store.load("chat"); const changed = session(); changed.messages.push({ kind: "assistant", text: "other window" });
   w.rows.set("chat", { session: changed, revision: 2 }); await w.store.refresh();
   w.store.save(old!); await w.store.flush(); expect(w.store.pendingCount).toBe(0);
   expect(w.rows.get("chat")?.session.messages.at(-1)?.text).toBe("other window");
 });

 it("H13 keeps independent pending queues from two windows without replacing either", async () => {
   const w = fixture();
   const first = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice, "window-a");
   const second = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice, "window-b");
   first.save({ ...session(), id: "a" }); second.save({ ...session(), id: "b" });
   const recovered = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice, "recovered");
   expect(recovered.pendingExport().map(s => s.id).sort()).toEqual(["a", "b"]);
 });
 it("H33 keeps both offline queues when a duplicated pane inherits the same writer ID", () => {
   const w = fixture(); const a = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice, "copied-session");
   const b = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice, "copied-session");
   a.save({ ...session(), id: "one" }); b.save({ ...session(), id: "two" });
   const reopened = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice, "new-pane");
   expect(reopened.pendingExport().map(s => s.id).sort()).toEqual(["one", "two"]);
 });

 it("H13 preserves a failed save across reload and retries after reconnect", async () => {
  const w = fixture(); await w.store.connect();
  w.api.historySave.mockRejectedValueOnce(new Error("offline")); w.store.save(session());
  await expect(w.store.flush()).rejects.toThrow("offline"); expect(w.notice).toHaveBeenCalled();
  const reloaded = new DurableHistory(() => w.api, () => "excel", localStorage, w.notice);
  await reloaded.connect(); await reloaded.flush(); expect(w.rows.get("chat")?.session.messages[0]?.text).toBe("hello"); expect(reloaded.pendingCount).toBe(0);
 });
 it("H14 serializes in-flight snapshots and rebases only its own acknowledged successor", async () => {
  const w = fixture(); await w.store.connect();
  let complete!: () => void; const waiting = new Promise<void>(resolve => { complete = resolve; });
  const save = w.api.historySave.getMockImplementation()!;
  w.api.historySave.mockImplementationOnce(async input => { await waiting; return save(input); });
  w.store.save(session()); const flushing = w.store.flush(); await Promise.resolve();
  const next = session(); next.messages.push({ kind: "assistant", text: "answer" }); w.store.save(next); complete();
  await flushing; expect(w.rows.get("chat")?.revision).toBe(2); expect(w.rows.get("chat")?.session.messages).toHaveLength(2);
 });
 it("H30 discards an identical stale revision without replacing a newer server record", async () => {
  const w = fixture(); await w.store.connect(); w.rows.set("chat", { session: session(), revision: 5 });
  w.store.save(session()); await w.store.flush(); expect(w.store.pendingCount).toBe(0); expect(w.rows.get("chat")?.revision).toBe(5);
 });
 it("H12 reports blocked browser storage while preserving existing local data", async () => {
  const w = fixture(); await w.store.connect();
  vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => { throw new Error("QuotaExceededError"); });
  w.store.save(session()); expect(w.notice).toHaveBeenCalledWith(expect.stringContaining("not"));
  await w.store.flush(); expect(w.rows.has("chat")).toBe(true);
 });
 it("H36 does not rewrite a chat when only its timestamp changed", async () => {
   const w = fixture(); await w.store.connect(); const chat = session(); w.store.save(chat); await w.store.flush();
   expect(w.api.historySave).toHaveBeenCalledTimes(1);
   w.store.save({ ...chat, updatedAt: chat.updatedAt + 5000 }); await w.store.flush();
   expect(w.api.historySave).toHaveBeenCalledTimes(1); expect(w.store.pendingCount).toBe(0);
   w.store.save({ ...chat, updatedAt: chat.updatedAt + 9000, messages: [...chat.messages, { kind: "assistant", text: "more" }] }); await w.store.flush();
   expect(w.api.historySave).toHaveBeenCalledTimes(2); expect(w.rows.get("chat")?.session.messages).toHaveLength(2);
 });
 it("H37 cuts an oversized tool result before saving, without touching the live chat", async () => {
   const w = fixture(); await w.store.connect(); const chat = session();
   chat.messages.push({ kind: "tool_group", text: "", tools: [{ name: "read_range", args: "{}", result: "r".repeat(MAX_STORED_TEXT * 3), state: "ok", error: null }] });
   w.store.save(chat); await w.store.flush();
   expect(w.rows.get("chat")?.session.messages[1]?.tools?.[0]?.result).toHaveLength(MAX_STORED_TEXT + 1);
   expect(chat.messages[1]?.tools?.[0]?.result).toHaveLength(MAX_STORED_TEXT * 3);
 });
 it("H38 keeps a chat that is too large to save without retrying it forever", async () => {
   vi.useFakeTimers(); const w = fixture(); await w.store.connect();
   w.api.historySave.mockRejectedValue(Object.assign(new Error("too large"), { code: "HISTORY_TOO_LARGE" }));
   w.store.save(session()); await expect(w.store.flush()).rejects.toThrow("too large");
   await vi.advanceTimersByTimeAsync(60_000);
   expect(w.api.historySave).toHaveBeenCalledTimes(1); expect(w.store.pendingCount).toBe(1);
   expect(w.notice.mock.calls.some(([m]) => String(m).includes("too large to save"))).toBe(true);
 });
 it("H39 reconnects by itself after the first connection fails", async () => {
   vi.useFakeTimers(); const w = fixture(); w.api.historyStats.mockRejectedValueOnce(new Error("offline"));
   await expect(w.store.connect()).rejects.toThrow("offline");
   await vi.advanceTimersByTimeAsync(1000); expect(w.api.historyStats).toHaveBeenCalledTimes(2);
   w.store.save(session()); await w.store.flush(); expect(w.rows.has("chat")).toBe(true);
 });
 it("H40 offers the old browser copy for export only until it has been imported", async () => {
   localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([session()]));
   const w = fixture(); expect(w.store.legacyUnmigrated()).not.toBeNull();
   await w.store.connect(); expect(w.store.legacyUnmigrated()).toBeNull();
 });
});
