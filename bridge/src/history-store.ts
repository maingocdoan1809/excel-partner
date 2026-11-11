import fs from "node:fs";
import path from "node:path";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
import type Database from "better-sqlite3";
import { StoredSessionSchema, type StoredSession, type SessionHost, type HistorySave, type HistoryAck, type HistoryRecord, type HistorySummary } from "../../protocol/src/history.js";

interface Row { host: SessionHost; id: string; metadata: string; revision: number; deleted: number; operation: string; hash: string; updated: number }
export class HistoryError extends Error { constructor(readonly code: string) { super(code); } }
type OpenDatabase = (file: string) => Database.Database;
/**
 * The SQLite driver is a native module built for specific Node.js versions. Load it only when History is first used,
 * so a missing or mismatched binary disables History instead of preventing the whole bridge from starting.
 */
const openNative: OpenDatabase = file => {
  const Driver = createRequire(import.meta.url)("better-sqlite3") as typeof Database;
  return new Driver(file, { timeout: 5000 });
};
const driverFailure = (error: unknown) => /better_sqlite3\.node|bindings file|ERR_DLOPEN_FAILED|NODE_MODULE_VERSION|Cannot find module 'better-sqlite3'/i.test(String((error as { code?: string })?.code ?? "") + String((error as Error)?.message ?? error));
/** One SQLite owner per bridge; transactions/CAS also protect against another process. */
export class HistoryStore {
  private db: Database.Database;
  private closed = false;
  constructor(readonly file: string, open: OpenDatabase = openNative) {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    try { this.db = open(file); } catch (error) { throw driverFailure(error) ? new HistoryError("HISTORY_DRIVER_UNAVAILABLE") : error; }
    try {
      const version = this.db.pragma("user_version", { simple: true }) as number;
      if (version > 1) throw new HistoryError("HISTORY_SCHEMA_NEWER");
      if (this.db.pragma("quick_check", { simple: true }) !== "ok") throw new HistoryError("HISTORY_CORRUPT");
      this.db.pragma("journal_mode = WAL"); this.db.pragma("synchronous = FULL"); this.db.pragma("foreign_keys = ON");
      this.db.transaction(() => {
        this.db.exec(`CREATE TABLE IF NOT EXISTS history_scope(host TEXT PRIMARY KEY, generation INTEGER NOT NULL DEFAULT 0);
          CREATE TABLE IF NOT EXISTS history_chats(host TEXT NOT NULL,id TEXT NOT NULL,metadata TEXT NOT NULL,updated REAL NOT NULL,revision INTEGER NOT NULL,deleted INTEGER NOT NULL DEFAULT 0,operation TEXT NOT NULL,hash TEXT NOT NULL,PRIMARY KEY(host,id));
          CREATE INDEX IF NOT EXISTS history_order ON history_chats(host,deleted,updated DESC,id);
          CREATE TABLE IF NOT EXISTS history_entries(host TEXT NOT NULL,chat TEXT NOT NULL,kind TEXT NOT NULL,ordinal INTEGER NOT NULL,data TEXT NOT NULL,PRIMARY KEY(host,chat,kind,ordinal),FOREIGN KEY(host,chat) REFERENCES history_chats(host,id));
          PRAGMA user_version = 1;`);
        for (const host of ["excel", "word", "powerpoint"]) this.db.prepare("INSERT OR IGNORE INTO history_scope(host) VALUES(?)").run(host);
      })();
      try { fs.chmodSync(file, 0o600); } catch { /* Windows ACLs apply. */ }
    } catch (error) { this.db.close(); throw error; }
  }
  close() { if (!this.closed) { this.closed = true; this.db.close(); } }
  private generation(host: SessionHost) { return (this.db.prepare("SELECT generation FROM history_scope WHERE host=?").get(host) as { generation: number }).generation; }
  private row(host: SessionHost, id: string) { return this.db.prepare("SELECT * FROM history_chats WHERE host=? AND id=?").get(host, id) as Row | undefined; }
  private entries(host: SessionHost, id: string, kind: string, offset = 0, limit = -1) {
    return (this.db.prepare("SELECT data FROM history_entries WHERE host=? AND chat=? AND kind=? ORDER BY ordinal LIMIT ? OFFSET ?").all(host, id, kind, limit, offset) as { data: string }[]).map(row => JSON.parse(row.data));
  }
  private count(host: SessionHost, id: string, kind: string) { return (this.db.prepare("SELECT count(*) n FROM history_entries WHERE host=? AND chat=? AND kind=?").get(host, id, kind) as { n: number }).n; }
  save(host: SessionHost, input: HistorySave): HistoryAck {
    const session = { ...StoredSessionSchema.parse(input.session), host } as StoredSession;
    const hash = createHash("sha256").update(JSON.stringify(session)).digest("hex");
    return this.db.transaction(() => {
      const generation = this.generation(host);
      if (input.generation !== generation) throw new HistoryError("HISTORY_GENERATION");
      const row = this.row(host, session.id);
      if (row?.deleted) throw new HistoryError("HISTORY_DELETED");
      if (row?.operation === input.operationId) {
        if (row.hash !== hash) throw new HistoryError("HISTORY_CONFLICT");
        return { revision: row.revision, generation };
      }
      if ((row?.revision ?? 0) !== input.baseRevision) throw new HistoryError("HISTORY_CONFLICT");
      const revision = (row?.revision ?? 0) + 1;
      const { messages, changes, ...metadata } = session;
      this.db.prepare(`INSERT INTO history_chats VALUES(?,?,?,?,?,0,?,?) ON CONFLICT(host,id) DO UPDATE SET metadata=excluded.metadata,updated=excluded.updated,revision=excluded.revision,operation=excluded.operation,hash=excluded.hash`).run(host, session.id, JSON.stringify(metadata), session.updatedAt, revision, input.operationId, hash);
      const upsert = this.db.prepare(`INSERT INTO history_entries VALUES(?,?,?,?,?) ON CONFLICT(host,chat,kind,ordinal) DO UPDATE SET data=excluded.data WHERE data<>excluded.data`);
      for (const [kind, items] of [["message", messages], ["change", changes]] as const) {
        items.forEach((item, index) => upsert.run(host, session.id, kind, index, JSON.stringify(item)));
        this.db.prepare("DELETE FROM history_entries WHERE host=? AND chat=? AND kind=? AND ordinal>=?").run(host, session.id, kind, items.length);
      }
      return { revision, generation };
    })();
  }
  list(host: SessionHost, offset = 0, limit = 40) { return this.db.transaction(() => this.listSnapshot(host, offset, limit))(); }
  private listSnapshot(host: SessionHost, offset = 0, limit = 40) {
    const items = (this.db.prepare("SELECT * FROM history_chats WHERE host=? AND deleted=0 ORDER BY updated DESC,id LIMIT ? OFFSET ?").all(host, Math.min(100, Math.max(1, limit)), Math.max(0, offset)) as Row[]).map(row => {
      const s = JSON.parse(row.metadata) as StoredSession;
      return { id: s.id, host, title: s.title, provider: s.provider, providerSessionId: s.providerSessionId, startedAt: s.startedAt, updatedAt: s.updatedAt, revision: row.revision,
        messageCount: (this.db.prepare("SELECT count(*) n FROM history_entries WHERE host=? AND chat=? AND kind='message' AND json_extract(data,'$.kind') IN ('user','assistant')").get(host, s.id) as { n: number }).n,
        changeCount: (this.db.prepare("SELECT count(*) n FROM history_entries WHERE host=? AND chat=? AND kind='change' AND json_extract(data,'$.ok')=1").get(host, s.id) as { n: number }).n } satisfies HistorySummary;
    });
    return { items, total: this.stats(host).count, generation: this.generation(host) };
  }
  get(host: SessionHost, id: string): HistoryRecord | null { return this.db.transaction(() => this.getSnapshot(host, id))(); }
  private getSnapshot(host: SessionHost, id: string): HistoryRecord | null {
    const row = this.row(host, id); if (!row || row.deleted) return null;
    return { session: { ...JSON.parse(row.metadata), messages: this.entries(host, id, "message"), changes: this.entries(host, id, "change") }, revision: row.revision, generation: this.generation(host) };
  }
  page(host: SessionHost, id: string, offset = 0, limit = 100) { return this.db.transaction(() => this.pageSnapshot(host, id, offset, limit))(); }
  private pageSnapshot(host: SessionHost, id: string, offset = 0, limit = 100) {
    const row = this.row(host, id); if (!row || row.deleted) return null;
    limit = Math.min(200, Math.max(1, limit)); offset = Math.max(0, offset);
    return { session: { ...JSON.parse(row.metadata), messages: this.entries(host, id, "message", offset, limit), changes: this.entries(host, id, "change", offset, limit) } as StoredSession,
      revision: row.revision, generation: this.generation(host), totalMessages: this.count(host, id, "message"), totalChanges: this.count(host, id, "change") };
  }
  delete(host: SessionHost, id: string) {
    this.db.transaction(() => {
      this.db.prepare("INSERT INTO history_chats VALUES(?,?,'{}',0,1,1,'','') ON CONFLICT(host,id) DO UPDATE SET deleted=1,revision=revision+1").run(host, id);
      this.db.prepare("DELETE FROM history_entries WHERE host=? AND chat=?").run(host, id);
    })();
  }
  clear(host: SessionHost) { this.db.transaction(() => {
    this.db.prepare("UPDATE history_scope SET generation=generation+1 WHERE host=?").run(host);
    this.db.prepare("UPDATE history_chats SET deleted=1,revision=revision+1 WHERE host=?").run(host);
    this.db.prepare("DELETE FROM history_entries WHERE host=?").run(host);
  })(); }
  importLegacy(sessions: StoredSession[]) {
    const parsed = sessions.map(session => StoredSessionSchema.parse(session) as StoredSession);
    let imported = 0;
    this.db.transaction(() => { for (const session of parsed) {
      const host = session.host ?? "excel";
      if (this.row(host, session.id) || this.generation(host) !== 0) continue;
      this.save(host, { session, baseRevision: 0, generation: 0, operationId: "legacy:" + createHash("sha256").update(session.id).digest("hex") }); imported++;
    } })(); return imported;
  }
  export(host: SessionHost) { return this.db.transaction(() => (this.db.prepare("SELECT id FROM history_chats WHERE host=? AND deleted=0 ORDER BY updated,id").all(host) as { id: string }[]).map(row => this.get(host, row.id)!.session))(); }
  stats(host: SessionHost) { return { count: (this.db.prepare("SELECT count(*) n FROM history_chats WHERE host=? AND deleted=0").get(host) as { n: number }).n,
    bytes: [this.file, this.file + "-wal", this.file + "-shm"].reduce((sum, file) => { try { return sum + fs.statSync(file).size; } catch { return sum; } }, 0), generation: this.generation(host) }; }
  cleanup(host: SessionHost, before: number, protectedIds: string[]) {
    return this.db.transaction(() => { const rows = this.db.prepare("SELECT id FROM history_chats WHERE host=? AND deleted=0 AND updated<?").all(host, before) as { id: string }[];
      const targets = rows.filter(row => !protectedIds.includes(row.id)); for (const row of targets) this.delete(host, row.id); return targets.length; })();
  }
}
