/** Isolated native SQLite contract check. No services, provider login or Office writes. */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import assert from "node:assert/strict";
import { HistoryStore } from "../dist/bridge/src/history-store.js";
if (process.platform !== "win32") throw new Error("Run this check with Windows Node.js and Windows-installed dependencies.");
const directory = fs.mkdtempSync(path.join(os.tmpdir(), "office-history-check-"));
let store;
try {
 const file = path.join(directory, "history.db");
 store = new HistoryStore(file);
 const session = { id: "sqlite-smoke", host: "excel", provider: "copilot", title: "SQLite smoke", startedAt: 1, updatedAt: 2,
   messages: [{ kind: "user", text: "WINDOWS_SQLITE_OK" }], changes: [] };
 store.save("excel", { session, baseRevision: 0, generation: 0, operationId: "first" });
 assert.equal(store.get("excel", session.id).session.messages[0].text, "WINDOWS_SQLITE_OK");
 assert.throws(() => store.save("excel", { session, baseRevision: 0, generation: 0, operationId: "stale" }), /HISTORY_CONFLICT/);
 store.close(); store = new HistoryStore(file); assert.equal(store.list("excel").total, 1);
 console.log(JSON.stringify({ node: process.versions.node, platform: process.platform, architecture: process.arch, persistence: true, conflict: true }));
} finally { store?.close(); fs.rmSync(directory, { recursive: true, force: true }); }
