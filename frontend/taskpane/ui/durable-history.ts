import type { HistoryAck, HistoryPage, HistorySave, HistoryStats, HistoryTranscript, SessionHost, SessionSummary, StoredSession } from "../../../protocol/src/history.js";
import { clipSession, inferHost, titleFrom } from "./sessions.js";
import { cleanSessionTitle } from "../../../protocol/src/session-title.js";
import { historyPrefix } from "./history-reconcile.js";
interface HistoryApi {
 historyList(offset?: number, limit?: number): Promise<HistoryPage>;
 historyPage(id: string, offset?: number, limit?: number): Promise<HistoryTranscript>;
 historyStats(): Promise<HistoryStats>;
 historySave(input: HistorySave, keepalive?: boolean): Promise<HistoryAck>;
 historyImport(sessions: StoredSession[]): Promise<{ imported: number }>;
 historyDelete(id?: string): Promise<HistoryStats>;
}
export interface ChatHistoryStore {
 list(): SessionSummary[]; get(id: string): StoredSession | null; save(session: StoredSession): void;
 delete(id: string): void | Promise<void>; clear(): void | Promise<void>;
 latest(maxAgeMs: number): StoredSession | null;
 loadPage?(id: string, offset?: number, limit?: number): Promise<HistoryTranscript>;
 refresh?(more?: boolean): Promise<void>; load?(id: string): Promise<StoredSession | null>; flush?(keepalive?: boolean): Promise<void>;
 total?: number;
 savedId?(id: string): string;
}
interface Pending extends HistorySave { host: SessionHost; recoveryId?: string }
const OUTBOX = "excelPartner.historyOutbox.v1";
/** How long a request waits for a bridge that answers HISTORY_RESTARTING (it is handing over to a new release). */
export const HISTORY_RESTART_WAIT_MS = 30_000;
const restarting = (error: unknown) => (error as { code?: string })?.code === "HISTORY_RESTARTING";
/** UI cache + acknowledged SQLite writes. localStorage is only migration input/outbox. */
export class DurableHistory implements ChatHistoryStore {
 private records = new Map<string, StoredSession>();
 private summaries: SessionSummary[] = [];
 private revisions = new Map<string, number>();
 private queue: Pending[] = [];
 private running?: Promise<void>;
 private inFlight?: Pending;
 private generation = 0;
 private ready = false;
 private listed = false;
 private migrationNotice?: string;
 private connecting?: Promise<void>;
 private readonly outboxKey: string;
 private readonly aliasKey: string;
 private aliases = new Map<string, string>();
 private recoveryIds = new Map<string, string>();
 // Content of the newest copy of each chat that was queued, without its timestamp: an unchanged chat is not saved again.
 private signatures = new Map<string, string>();
 private sources = new Map<string, Set<string>>();
 private retryTimer?: ReturnType<typeof setTimeout>;
 private retryDelay = 1000;
 private failureNotices = new Set<string>();
 total = 0;
 /** The bridge's history API, waiting quietly through a restart instead of reporting it as a failure. */
 private readonly api: HistoryApi = {
   historyList: (...args) => this.retrying(() => this.client().historyList(...args)),
   historyPage: (...args) => this.retrying(() => this.client().historyPage(...args)),
   historyStats: () => this.retrying(() => this.client().historyStats()),
   historySave: (...args) => this.retrying(() => this.client().historySave(...args)),
   historyImport: (...args) => this.retrying(() => this.client().historyImport(...args)),
   historyDelete: (...args) => this.retrying(() => this.client().historyDelete(...args)),
 };
 private async retrying<T>(request: () => Promise<T>): Promise<T> {
   for (let waited = 0; ; waited += 1000) {
     try { return await request(); }
     catch (error) { if (!restarting(error) || waited >= HISTORY_RESTART_WAIT_MS) throw error; }
     await new Promise(resolve => setTimeout(resolve, 1000));
   }
 }
 constructor(private readonly client: () => HistoryApi, private readonly host: () => SessionHost, private readonly storage: Pick<Storage, "getItem" | "setItem" | "removeItem">,
   private readonly notify: (message: string) => void = () => {}, writerId?: string) {
   if (!writerId) { try { writerId = sessionStorage.getItem("excelPartner.historyWriter") ?? crypto.randomUUID(); sessionStorage.setItem("excelPartner.historyWriter", writerId); } catch { writerId = crypto.randomUUID(); } }
   // Duplicated browser panes can inherit sessionStorage. Each running instance
   // still needs its own outbox key so offline writes cannot replace each other.
   this.outboxKey = `${OUTBOX}.${writerId}.${crypto.randomUUID()}`;
   this.aliasKey = `excelPartner.historyAliases.v1.${writerId}`;
   try { const raw = JSON.parse(storage.getItem(this.aliasKey) ?? "{}"); for (const [key, id] of Object.entries(raw)) if (typeof id === "string" && id.length <= 256) this.aliases.set(key, id); } catch { /* Outbox still preserves unsaved content. */ }
   try {
     const source = JSON.parse(storage.getItem("excelPartner.sessions.v1") ?? "[]") as StoredSession[];
     if (Array.isArray(source)) for (const s of source) if (s && typeof s.id === "string" && Array.isArray(s.messages)) this.records.set(`${s.host ?? inferHost(s)}:${s.id}`, { ...s, host: s.host ?? inferHost(s) });
   } catch { this.notify("Old browser history could not be read. Its original bytes have been retained for recovery."); }
   try {
     const indexed = storage as Storage;
     const keys = new Set([OUTBOX, this.outboxKey]);
     if (typeof indexed.key === "function") for (let i = 0; i < indexed.length; i++) { const key = indexed.key(i); if (key?.startsWith(OUTBOX + ".")) keys.add(key); }
     const pending = new Map<string, Pending>();
     for (const key of keys) {
       const queued = JSON.parse(storage.getItem(key) ?? "[]") as Pending[];
       if (!Array.isArray(queued)) continue;
       for (const item of queued) if (item?.session?.id && Array.isArray(item.session.messages) && ["excel", "word", "powerpoint"].includes(item.host)) {
         const identity = `${item.host}:${item.session.id}`;
         item.recoveryId ??= this.recoveryIds.get(identity) ?? `recovered-${item.operationId}`;
         this.recoveryIds.set(identity, item.recoveryId);
         pending.set(item.operationId, item); const sources = this.sources.get(item.operationId) ?? new Set(); sources.add(key); this.sources.set(item.operationId, sources);
       }
     }
     this.queue = [...pending.values()];
     for (const item of this.queue) if (item.host === this.host()) this.records.set(`${item.host}:${item.session.id}`, item.session);
   } catch { this.notify("History recovery data could not be read. Export the existing browser data before retrying migration."); }
 }
 get pendingCount() { return this.queue.filter(item => item.host === this.host()).length; }
 private outbox() {
   try {
     const pending = this.queue.filter(item => item.host === this.host());
     if (pending.length) this.storage.setItem(this.outboxKey, JSON.stringify(pending)); else this.storage.removeItem(this.outboxKey);
   } catch { this.notify("Browser storage is full or blocked. Pending history is not protected against closing this pane; keep it open until saved or export the pending copy."); }
 }
 async connect() {
   if (this.connecting) return this.connecting;
   return this.connecting = this.initialize().finally(() => { this.connecting = undefined; });
 }
 private async initialize() {
   try {
     const host = this.host(), marker = `excelPartner.historyMigrated.v1.${host}`;
     const stats = await this.api.historyStats(); this.generation = stats.generation;
     let legacy: string | null = null, migrated = false, readable = true;
     try { migrated = this.storage.getItem(marker) === "1"; legacy = this.storage.getItem("excelPartner.sessions.v1"); }
     catch { readable = false; this.notify("Old browser history could not be read. New chats can still be saved on this computer; retry migration when browser storage is available."); }
     this.migrationNotice = undefined;
     if (!migrated && readable) {
       try {
       const raw = JSON.parse(legacy ?? "[]") as StoredSession[];
       if (!Array.isArray(raw)) throw new Error("Invalid legacy history");
       const source = raw.filter(s => (s.host ?? inferHost(s)) === host).map(s => ({ ...s, host }));
       for (let i = 0; i < source.length; i += 10) await this.api.historyImport(source.slice(i, i + 10));
       // Keep legacy bytes for old retained app releases; only mark an acknowledged import.
       try { this.storage.setItem(marker, "1"); } catch { this.notify("History imported, but the migration marker could not be saved. A retry is safe."); }
       } catch { this.migrationNotice = "Old browser history could not be imported. New chats can still be saved; the old browser data is retained. Use Export history in Settings to keep a copy."; this.notify(this.migrationNotice); }
     }
     await this.refresh(); this.ready = true;
     // A write failure must not turn a successfully loaded History into "unavailable".
     await this.flush().catch(() => {});
     this.retryDelay = 1000;
   } catch (error) {
     this.notify(`History is not fully saved or available: ${String(error instanceof Error ? error.message : error)}`);
     // Nothing else would try again until the pane reconnects, so keep trying with a growing pause.
     if (!this.retryTimer) {
       this.retryTimer = setTimeout(() => { this.retryTimer = undefined; void this.connect().catch(() => {}); }, this.retryDelay);
       this.retryDelay = Math.min(30_000, this.retryDelay * 2);
     }
     throw error;
   }
 }
 list(): SessionSummary[] {
   if (this.listed) return this.summaries;
   return [...this.records.values()].filter(s => s.host === this.host() && s.messages.some(m => m.kind === "user")).sort((a,b) => b.updatedAt-a.updatedAt).map(s => this.summary(s));
 }
 private summary(s: StoredSession): SessionSummary { return { id: s.id, title: s.title, provider: s.provider, providerSessionId: s.providerSessionId, startedAt: s.startedAt, updatedAt: s.updatedAt, messageCount: s.messages.filter(m => m.kind === "user" || m.kind === "assistant").length, changeCount: s.changes.filter(c => c.ok).length }; }
 private storageId(id: string) {
   const seen = new Set<string>();
   while (this.aliases.has(`${this.host()}:${id}`) && !seen.has(id)) { seen.add(id); id = this.aliases.get(`${this.host()}:${id}`)!; }
   return id;
 }
 savedId(id: string) { return this.storageId(id); }
 get(id: string) { const s = this.records.get(`${this.host()}:${id}`); return s?.host === this.host() ? structuredClone(s) : null; }
 getActive(id: string) { return this.get(this.storageId(id)); }
 loadActive(id: string) { return this.load(this.storageId(id)); }
 private saveAliases() { try { if (this.aliases.size) this.storage.setItem(this.aliasKey, JSON.stringify(Object.fromEntries(this.aliases))); else this.storage.removeItem(this.aliasKey); } catch { /* Pending snapshots retain their resolved IDs. */ } }
 latest(maxAgeMs: number) { const s = [...this.records.values()].filter(s => s.host === this.host()).sort((a,b) => b.updatedAt-a.updatedAt)[0]; return s && Date.now()-s.updatedAt<=maxAgeMs ? structuredClone(s) : null; }
 async refresh(more = false) {
   const page = await this.api.historyList(more ? this.summaries.length : 0, 40);
   if (!Array.isArray(page.items)) throw new Error("Invalid history response");
   this.summaries = more ? [...this.summaries, ...page.items.filter(s => !this.summaries.some(old => old.id === s.id))] : page.items;
   this.listed = true;
   this.total = page.total; this.generation = page.generation;
   // Summary refreshes cannot advance the revision of an older loaded transcript.
 }
 loadPage(id: string, offset = 0, limit = 100) { return this.api.historyPage(id, offset, limit); }
 async load(id: string) {
   const pending = this.queue.filter(p => p.host === this.host() && p.session.id === id).slice(-1)[0];
   if (pending) return structuredClone(pending.session);
   let first;
   for (let attempt = 0; attempt < 3; attempt++) {
     try { first = await this.readSaved(id); break; } catch (error) { if (!this.conflict(error) || attempt === 2) throw error; }
   }
   if (!first) throw new Error("History is still updating. Try opening this chat again.");
   const session = first.session;
   // Explicitly continuing a saved row selects that branch. Pane reloads use loadActive.
   if (this.aliases.delete(`${this.host()}:${id}`)) this.saveAliases();
   this.records.set(`${this.host()}:${id}`, session); this.revisions.set(id, first.revision); this.generation = first.generation;
   this.signatures.set(`${this.host()}:${id}`, this.signature(this.prepare(session)));
   return structuredClone(session);
 }
 private prepare(session: StoredSession): StoredSession {
   return clipSession(structuredClone({ ...session, host: this.host(), title: cleanSessionTitle(session.providerTitle) || titleFrom(session.messages) }));
 }
 private signature(session: StoredSession) { return JSON.stringify({ ...session, updatedAt: 0 }); }
 save(session: StoredSession) {
   const originalId = session.id, id = this.storageId(originalId);
   if (!session.messages.some(m => m.kind === "user") && !this.records.has(`${this.host()}:${session.id}`)) return;
   const copy = this.prepare(session);
   if (id !== originalId || id.startsWith("recovered-")) { copy.id = id; copy.providerSessionId = undefined; copy.providerTitle = undefined; copy.identityVersion = 1; copy.title += " (saved copy)"; }
   const key = `${this.host()}:${copy.id}`, signature = this.signature(copy);
   // Opening Settings or History, or hiding the pane, asks for a save without changing the chat. Do not rewrite it.
   if (this.signatures.get(key) === signature) return;
   this.signatures.set(key, signature);
   this.records.set(key, copy);
   const last = this.queue.filter(p => p.host === this.host() && p.session.id === copy.id).slice(-1)[0];
   if (last && last !== this.inFlight && !this.sources.has(last.operationId)) { last.session = copy; last.operationId = crypto.randomUUID(); }
   else {
     const key = `${this.host()}:${copy.id}`, recoveryId = last?.recoveryId ?? this.recoveryIds.get(key) ?? `recovered-${crypto.randomUUID()}`;
     this.recoveryIds.set(key, recoveryId);
     this.queue.push({ host: this.host(), session: copy, baseRevision: last ? last.baseRevision + 1 : this.revisions.get(copy.id) ?? 0, generation: last?.generation ?? this.generation, operationId: crypto.randomUUID(), recoveryId });
   }
   this.outbox();
   if (this.ready && !this.retryTimer) this.retryTimer = setTimeout(() => { this.retryTimer = undefined; void this.flush().catch(() => {}); }, 250);
 }
 async flush(keepalive = false): Promise<void> {
   if (this.retryTimer) { clearTimeout(this.retryTimer); this.retryTimer = undefined; }
   if (this.running) return this.running;
   if (!this.ready) return;
   return this.running = this.drain(keepalive).finally(() => { this.running = undefined; this.inFlight = undefined; });
 }
 private async drain(keepalive: boolean) {
   let saved = false, failure: unknown;
   const blocked = new Set<string>();
   while (true) {
     const next = this.queue.find(p => p.host === this.host() && !blocked.has(p.session.id)); if (!next) break;
     this.inFlight = next;
     try {
       const deleting = !next.session.messages.some(m => m.kind === "user");
       const result = deleting ? { ack: { revision: 0, ...(await this.api.historyDelete(next.session.id)) }, advance: true } : await this.saveReconciled(next, keepalive);
       const ack = result.ack;
       if (deleting) { this.records.delete(`${this.host()}:${next.session.id}`); this.signatures.delete(`${this.host()}:${next.session.id}`); }
       this.forgetRecovery(next.operationId);
       if (result.advance) this.revisions.set(next.session.id, ack.revision);
       this.queue.splice(this.queue.indexOf(next), 1);
       const successor = this.queue.find(p => p.host === next.host && p.session.id === next.session.id);
       if (successor && result.advance) successor.baseRevision = ack.revision;
       this.outbox(); saved = true;
     } catch (error) {
       failure = error; blocked.add(next.session.id);
       const code = (error as { code?: string })?.code;
       const removed = code === "HISTORY_DELETED" || code === "HISTORY_GENERATION";
       const tooLarge = code === "HISTORY_TOO_LARGE";
       const key = `${next.session.id}:${code ?? String(error)}`;
       if (!this.failureNotices.has(key)) {
         this.failureNotices.add(key);
         this.notify(removed ? "A pending chat was removed in another window. Its local copy is kept; the rest of History is available."
           : tooLarge ? "This chat is too large to save. Export what you need, then start a new chat."
           : this.conflict(error) ? "History is updating in another window. Your changes are queued and saving will retry automatically."
           : `History could not be saved yet. Your pending copy is kept. ${error instanceof Error ? error.message : String(error)}`);
       }
       if (!removed && !tooLarge && !keepalive && !this.retryTimer) {
         this.retryTimer = setTimeout(() => { this.retryTimer = undefined; void this.flush().catch(() => {}); }, this.retryDelay);
         this.retryDelay = Math.min(30_000, this.retryDelay * 2);
       }
     }
   }
   if (saved) await this.refresh();
   if (failure) throw failure;
   if (saved) { this.retryDelay = 1000; this.failureNotices.clear(); this.notify(this.migrationNotice ?? "History saved on this computer."); }
 }
 private conflict(error: unknown) { return (error as { code?: string })?.code === "HISTORY_CONFLICT" || error instanceof Error && error.message === "HISTORY_CONFLICT"; }
 private async readSaved(id: string) {
   const first = await this.api.historyPage(id, 0, 100);
   const session = structuredClone(first.session);
   for (let offset = 100; offset < Math.max(first.totalMessages, first.totalChanges); offset += 100) {
     const page = await this.api.historyPage(id, offset, 100);
     if (page.revision !== first.revision) throw Object.assign(new Error("HISTORY_CONFLICT"), { code: "HISTORY_CONFLICT" });
     session.messages.push(...page.session.messages); session.changes.push(...page.session.changes);
   }
   return { ...first, session };
 }
 private async saveReconciled(next: Pending, keepalive: boolean): Promise<{ ack: HistoryAck; advance: boolean }> {
   let error: unknown;
   for (let attempt = 0; attempt < 4; attempt++) {
     try { return { ack: await this.api.historySave(next, keepalive), advance: true }; }
     catch (failure) { if (!this.conflict(failure)) throw failure; error = failure; }
     let current;
     try { current = await this.readSaved(next.session.id); }
     catch (failure) { if (this.conflict(failure)) continue; throw failure; }
     const localExtends = historyPrefix(current.session, next.session);
     if (historyPrefix(next.session, current.session)) return { ack: { revision: current.revision, generation: current.generation }, advance: localExtends };
     if (localExtends) { next.baseRevision = current.revision; next.generation = current.generation; this.outbox(); continue; }
     // Different edits remain separate. Stable IDs make orphan-outbox retries idempotent.
     const originalId = next.session.id;
     let recoveredId = next.recoveryId === originalId ? `recovered-${next.operationId}` : next.recoveryId ?? `recovered-${next.operationId}`;
     if (recoveredId === originalId) recoveredId = `recovered-${crypto.randomUUID()}`;
     this.aliases.set(`${this.host()}:${originalId}`, recoveredId);
     this.saveAliases();
     for (const pending of this.queue) if (pending.host === this.host() && pending.session.id === originalId) {
       pending.session = { ...pending.session, id: recoveredId, providerSessionId: undefined, providerTitle: undefined, identityVersion: 1, title: pending.session.title + " (saved copy)" };
       pending.baseRevision = 0; pending.generation = current.generation; pending.recoveryId = recoveredId;
       this.records.set(`${this.host()}:${recoveredId}`, structuredClone(pending.session));
     }
     this.outbox();
   }
   throw error;
 }
 private forgetRecovery(operationId: string) {
   for (const key of this.sources.get(operationId) ?? []) {
     if (key === this.outboxKey) continue;
     try { const records = JSON.parse(this.storage.getItem(key) ?? "[]") as Pending[]; const rest = records.filter(item => item.operationId !== operationId); if (rest.length) this.storage.setItem(key, JSON.stringify(rest)); else this.storage.removeItem(key); }
     catch { this.notify("History saved, but a recovery copy could not be cleared. Retry is safe."); }
   }
   this.sources.delete(operationId);
 }
 private removeLegacy(id?: string) {
   try {
     const raw = JSON.parse(this.storage.getItem("excelPartner.sessions.v1") ?? "[]") as StoredSession[];
     const rest = raw.filter(s => (s.host ?? inferHost(s)) !== this.host() || id !== undefined && s.id !== id);
     this.storage.setItem("excelPartner.sessions.v1", JSON.stringify(rest));
   } catch { this.notify("Saved history was deleted, but an old browser backup could not be updated. Clear that backup before sharing this profile."); }
 }
 async delete(id: string) {
   await this.running?.catch(() => {}); const stats = await this.api.historyDelete(id); this.removeLegacy(id);
   for (const item of this.queue) if (item.host === this.host() && item.session.id === id) this.forgetRecovery(item.operationId);
   this.queue = this.queue.filter(p => p.host !== this.host() || p.session.id !== id); this.records.delete(`${this.host()}:${id}`); this.signatures.delete(`${this.host()}:${id}`); this.revisions.delete(id); this.outbox();
   this.summaries = this.summaries.filter(s => s.id !== id); this.total = stats.count;
   for (const [key, target] of this.aliases) if (key === `${this.host()}:${id}` || target === id) this.aliases.delete(key);
   for (const [key, target] of this.recoveryIds) if (key === `${this.host()}:${id}` || target === id) this.recoveryIds.delete(key);
   this.saveAliases();
   await this.refresh().catch(() => this.notify("Chat deleted. The list will refresh when the connection is available."));
 }
 async clear() {
   await this.running?.catch(() => {}); const stats = await this.api.historyDelete(); this.removeLegacy(); this.generation = stats.generation;
   for (const item of this.queue) if (item.host === this.host()) this.forgetRecovery(item.operationId);
   this.queue = this.queue.filter(p => p.host !== this.host()); for (const [id,s] of this.records) if (s.host === this.host()) this.records.delete(id);
   for (const key of this.signatures.keys()) if (key.startsWith(this.host() + ":")) this.signatures.delete(key);
   this.revisions.clear(); this.outbox(); this.summaries = []; this.total = stats.count;
   for (const key of this.aliases.keys()) if (key.startsWith(this.host() + ":")) this.aliases.delete(key);
   this.recoveryIds.clear(); this.saveAliases();
   await this.refresh().catch(() => this.notify("History deleted. The list will refresh when the connection is available."));
 }
 async recoverPending() {
   await this.running?.catch(() => {});
   const stats = await this.api.historyStats();
   const latest = new Map<string, Pending>(); for (const item of this.queue) if (item.host === this.host()) latest.set(item.session.id, item);
   const recovered: Array<{ originalId: string; session: StoredSession }> = [];
   for (const [originalId, item] of latest) {
     if (!item.session.messages.some(m => m.kind === "user")) continue;
     const session = { ...structuredClone(item.session), id: `recovered-${crypto.randomUUID()}`, providerSessionId: undefined, providerTitle: undefined, identityVersion: 1 as const, updatedAt: Date.now() };
     const ack = await this.api.historySave({ session, baseRevision: 0, generation: stats.generation, operationId: crypto.randomUUID() });
     for (const old of this.queue) if (old.host === this.host() && old.session.id === originalId) this.forgetRecovery(old.operationId);
     this.queue = this.queue.filter(old => old.host !== this.host() || old.session.id !== originalId);
     this.signatures.delete(`${this.host()}:${originalId}`);
     this.records.set(`${this.host()}:${session.id}`, session); this.revisions.set(session.id, ack.revision); this.generation = ack.generation; this.outbox();
     recovered.push({ originalId, session });
   }
   await this.refresh(); return recovered;
 }
 pendingExport() { return this.queue.filter(p => p.host === this.host()).map(p => p.session); }
 /** The old browser copy of this app's chats, while it has not been imported into saved history; otherwise null. */
 legacyUnmigrated(): string | null {
   try { return this.storage.getItem(`excelPartner.historyMigrated.v1.${this.host()}`) === "1" ? null : this.storage.getItem("excelPartner.sessions.v1"); } catch { return null; }
 }
}
