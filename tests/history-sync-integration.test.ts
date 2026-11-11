// @vitest-environment jsdom
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, expect, it, vi } from "vitest";
import { DurableHistory } from "../addin/taskpane/ui/durable-history.js";
import { HistoryStore } from "../bridge/src/history-store.js";
import type { HistorySave, StoredSession } from "../protocol/src/history.js";
const stores: HistoryStore[] = [], dirs: string[] = [];
afterEach(() => { localStorage.clear(); stores.splice(0).forEach(s => s.close()); dirs.splice(0).forEach(d => fs.rmSync(d, { recursive: true, force: true })); });
const chat = (): StoredSession => ({ id: "copilot:native", providerSessionId: "native", provider: "copilot", host: "excel", title: "Old chat", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "original" }], changes: [] });
function fixture() {
 const dir = fs.mkdtempSync(path.join(os.tmpdir(), "history-sync-")); dirs.push(dir); const db = new HistoryStore(path.join(dir, "history.db")); stores.push(db);
 const api = { historyStats: async () => db.stats("excel"), historyList: async (offset = 0, limit = 40) => db.list("excel", offset, limit),
 historyPage: async (id: string, offset = 0, limit = 100) => { const page = db.page("excel", id, offset, limit); if (!page) throw Object.assign(Error("not found"), { code: "HISTORY_NOT_FOUND" }); return page; },
 historySave: async (input: HistorySave) => db.save("excel", input), historyImport: async (sessions: StoredSession[]) => ({ imported: db.importLegacy(sessions) }),
 historyDelete: async (id?: string) => { if (id) db.delete("excel", id); else db.clear("excel"); return db.stats("excel"); } };
 const pane = (id: string) => new DurableHistory(() => api, () => "excel", localStorage, vi.fn(), id);
 return { api, db, pane };
}
it("H29 migrates and continues a large legacy transcript against real SQLite without manual recovery", async () => {
 const w = fixture(), old = chat(); old.messages = Array.from({ length: 251 }, (_, i) => ({ kind: "user", text: `old ${i}` }));
 localStorage.setItem("excelPartner.sessions.v1", JSON.stringify([old])); const pane = w.pane("legacy");
 const continued = structuredClone(old); continued.messages.push({ kind: "assistant", text: "new answer" }); pane.save(continued);
 await pane.connect(); expect(pane.pendingCount).toBe(0); expect(w.db.list("excel").total).toBe(1);
 expect(w.db.get("excel", old.id)?.session.messages).toHaveLength(252);
 const reopened = w.pane("reopened"); await reopened.connect(); expect((await reopened.load(old.id))?.messages).toHaveLength(252);
});
it("H30/H32 reconciles two panes, preserves divergent edits and keeps one stable recovery copy", async () => {
 const w = fixture(); await w.api.historyImport([chat()]); const a = w.pane("a"), b = w.pane("b"); await a.connect(); await b.connect();
 const first = (await a.load(chat().id))!, second = (await b.load(chat().id))!;
 first.messages.push({ kind: "assistant", text: "answer from first pane" }); a.save(first); await a.flush();
 b.save(second); await b.flush(); expect(b.pendingCount).toBe(0); expect(w.db.list("excel").total).toBe(1);
 second.messages.push({ kind: "user", text: "different branch from second pane" }); b.save(second); await b.flush();
 const copy = w.db.list("excel").items.find(s => s.id !== first.id)!;
 expect(copy).toBeDefined(); expect(w.db.get("excel", first.id)?.session.messages.at(-1)?.text).toBe("answer from first pane");
 second.messages.push({ kind: "assistant", text: "second pane continuation" }); b.save(second); await b.flush();
 expect(w.db.list("excel").total).toBe(2); expect(w.db.get("excel", copy.id)?.session.messages).toHaveLength(3);
 const reloaded = w.pane("b"); await reloaded.connect(); expect((await reloaded.loadActive(second.id))?.id).toBe(copy.id);
 // Selecting the original saved row explicitly must not redirect back to the copy.
 expect((await reloaded.load(first.id))?.messages.at(-1)?.text).toBe("answer from first pane");
 expect(reloaded.savedId(first.id)).toBe(first.id);
 await b.delete(copy.id); expect(b.savedId(first.id)).toBe(first.id); expect(w.db.get("excel", copy.id)).toBeNull();
});
it("H07/H31 keeps deleted-chat pending data without resurrecting it or blocking History connection", async () => {
 const w = fixture(); await w.api.historyImport([chat()]); const pane = w.pane("deleted"); await pane.connect(); const old = (await pane.load(chat().id))!;
 w.db.delete("excel", old.id); old.messages.push({ kind: "assistant", text: "pending answer" }); pane.save(old);
 await expect(pane.connect()).resolves.toBeUndefined(); expect(pane.pendingCount).toBe(1);
 expect(w.db.get("excel", old.id)).toBeNull(); expect(pane.list()).toHaveLength(0);
});
it("H34 recognizes streamed text and tool completion while a follow-up message is queued", async () => {
 const w = fixture(), before = chat();
 before.messages.push({ kind: "tool_group", text: "", tools: [{ name: "read", args: "{}", state: "running", error: null }] },
   { kind: "assistant", text: "part" }, { kind: "user", text: "follow-up", delivery: "pending", queueId: "q" });
 await w.api.historyImport([before]); const pane = w.pane("stream"); await pane.connect();
 const after = structuredClone(before); after.messages[1]!.tools![0]!.state = "ok"; after.messages[1]!.tools![0]!.result = "done";
 after.messages[2]!.text = "partial answer completed"; after.messages[3]!.delivery = undefined;
 await w.api.historySave({ session: after, baseRevision: 1, generation: 0, operationId: "completed-in-another-pane" });
 pane.save(before); await pane.flush(); expect(pane.pendingCount).toBe(0); expect(w.db.list("excel").total).toBe(1);
 expect(w.db.get("excel", before.id)?.session.messages[2]?.text).toBe("partial answer completed");
});
