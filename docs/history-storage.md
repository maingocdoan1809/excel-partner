# Durable chat history

Office Partner stores chat history in SQLite at
`%LOCALAPPDATA%\Excel Partner\history.db`. The local bridge owns the database;
no database service, account, port or separate SQLite installation is needed.
Credentials remain managed by the provider CLIs. Gateway's official-client
conversations remain separate from Office Partner history.

## Saving and recovery

The bridge authenticates each history request using the connected pane's token,
and derives the Office host from that session. SQLite transactions, revisions,
idempotency keys and deletion tombstones prevent stale snapshots from overwriting
newer records or resurrecting deleted chats. Messages and change records are
stored separately; only changed entries are updated. Lists and transcript reads
are paginated. Storage does not silently prune full transcripts or old chats.
Model-context limits and UI display pages do not truncate stored history.

Browser localStorage holds preferences and a per-window recovery outbox. Writes
are queued while typing/streaming and flushed during a turn, at completion,
when hidden and best-effort on pagehide. The outbox survives normal reloads and
closing/reopening panes while browser storage remains available. The bridge
only acknowledges committed SQLite writes. A failed save is visible in the
pane and Settings. If both browser storage and the bridge are unavailable,
keep the pane open and export its pending copy; no system can promise crash
recovery for bytes that could not reach either storage medium.

Revision conflicts are reconciled automatically when transcript content is
compatible. Already saved older snapshots are removed from the outbox; longer
streamed transcripts are retried safely against the latest revision. Genuine
divergence is saved as a separate, stable “saved copy”, preserving the original.
Recovery copies clear native resume identity and never send a prompt on their
own. Ordinary migration, reopening and retrying stale outboxes require no export
or manual conflict resolution. Saved History stays readable if a write fails,
and temporary failures retry automatically with bounded backoff.

Explicitly deleted chats and old clear-history generations are not recreated.
Their unacknowledged pending contents are retained; the advanced recovery/export
controls remain available for these cases and persistent disk/permission errors.
Outbox keys are unique per running pane, including duplicated browser tabs.

History restores the exact per-pane active chat ID after a managed reload.
Ordinary reopening still starts a new chat. Undo remains memory-only; stored
change records do not restore the ability to undo prior document changes.

## Migration, update and rollback

On its first connection, each Office application imports its own legacy
`excelPartner.sessions.v1` records. Import is transactional per batch and
idempotent; already saved records and tombstones are not replaced. The old
browser bytes are retained and the completion marker is written only after
server acknowledgement. A failed import keeps the original data for retry.
Explicit delete/clear also removes matching legacy records after server success.

The database and backups live outside immutable app releases. App update,
rollback, repair and uninstall retain them. The SQLite schema is versioned;
an unsupported future schema or corrupt file fails closed and is not recreated.
Releases predating SQLite can display their retained browser snapshot; they do
not display newer SQLite-only chats. Updating forward restores access to the
retained database. Never downgrade or overwrite the database to make an older
application understand it.

## Storage controls

Settings > App & storage > Chat history provides:

- Saved-chat count, shared database/WAL size and count of chats not saved yet.
- One JSON export of this application's history: saved chats, plus any unsaved
  copies and old browser history that has not been imported yet.
- Recovery of unsaved chats as separate chats; the button appears only while
  chats are waiting to be saved. Saving is retried automatically.
- Explicit age-based cleanup scoped to this Office application, protecting
  current connected and pending chats. No automatic retention policy runs.

SQLite may retain freed pages for reuse after cleanup; deleting chat rows does
not guarantee immediate reduction of the database file's allocated size.
JSON export is host-scoped and includes chat contents, not provider
credential/config files. Keep the original database if diagnosis fails; to copy
it by hand, stop the bridge first and copy `history.db` with its `-wal` and
`-shm` files. Do not replace an open database.

Long fields are cut when a chat is saved (20,000 characters per message or tool
field, 200,000 for a generated document) so one large tool result cannot make a
chat too big to save. This only shortens the copy shown in History; a native
provider session that is resumed keeps its own full context. A chat that is
still too large is rejected with a clear message and is not retried.

If the SQLite driver cannot be loaded (for example an unsupported Node.js
version), only History is unavailable: the bridge and chats keep working and
History reports the problem.

## Windows packaging and compatibility

The exact driver pin is `better-sqlite3@12.9.0`. Its
[official release artifacts](https://github.com/WiseLibs/better-sqlite3/releases/tag/v12.9.0)
include Windows x64 and ARM64 binaries for supported Node ABIs, including Node 20.
Later driver releases omit Node 20 Windows binaries despite listing Node 20 in
package engines. This release therefore supports Node 20.19+ within major 20,
or Node 22.12+ below major 26. Node 26+ is explicitly outside this release's
range until a suitable driver/runtime pairing is selected. npm still needs
network access to fetch the prebuilt binary, as it does for other dependencies.

Actual isolated Windows x64 checks cover Node 20.19.0, 22.12.0 and 24.22.1:
compiled HistoryStore writes/reopens a database and rejects stale revisions
using downloaded prebuilt binaries, without invoking a compiler. This is not a pristine Windows VM or a Windows ARM64 execution test.

The patched build was also installed into an existing Windows installation on
Node 24.22.1. Chrome opened `https://localhost:8765/index.html` with normal HTTPS
verification, committed history to SQLite, rejected stale writes and read the
chat after reload. A real Copilot conversation completed and resumed the same
native identity after reload, with no tool calls. The bridge and gateway remain
healthy; validation removed only its own synthetic chat records. Real document
edits still require testing inside desktop Office.

Repeat in a Windows checkout with Windows dependencies after building:

```powershell
npm run typecheck
npm run build
npm test
node scripts/test-history-windows.mjs
npm run test:browser
```

The full acceptance matrix and execution record are kept in the repository
(`docs/runtime-history-testcases.md`) and are not part of the installed package.
Default tests and browser flows use mocks; real provider and Office document
checks are separate.
