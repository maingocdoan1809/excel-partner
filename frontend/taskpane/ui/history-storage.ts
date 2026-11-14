import type { BridgeClient } from "../bridge-client.js";
import type { DurableHistory } from "./durable-history.js";

/** Settings controls for saved history: one export, and recovery only while chats are unsaved. */
export function setupHistoryStorage(getClient: () => BridgeClient, store: DurableHistory,
  recover: (records: Awaited<ReturnType<DurableHistory["recoverPending"]>>) => void, isBusy: () => boolean) {
  const info = document.getElementById("history-storage-info");
  const status = document.getElementById("history-storage-status");
  const recoverButton = document.getElementById("history-recover-pending");
  const say = (text: string) => { if (status) status.textContent = text; };

  async function refresh() {
    try {
      const stats = await getClient().historyStats();
      const unsaved = store.pendingCount ? ` · ${store.pendingCount} not saved yet` : "";
      if (info) info.textContent = `${(stats.bytes / 1024 / 1024).toFixed(1)} MB on disk${unsaved}`;
    } catch {
      say("History storage is unavailable. Unsaved chats are kept and saved again automatically.");
    }
    if (recoverButton) recoverButton.hidden = store.pendingCount === 0;
  }

  function download(name: string, value: unknown) {
    const url = URL.createObjectURL(new Blob([JSON.stringify(value, null, 2)], { type: "application/json" }));
    const link = document.createElement("a");
    link.href = url; link.download = name; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  let busy = false;
  const on = (id: string, action: () => Promise<void | false>) => document.getElementById(id)?.addEventListener("click", () => {
    if (busy) return;
    busy = true;
    void action()
      .then(result => result === false ? undefined : refresh())
      .catch(error => say(error instanceof Error ? error.message : "History operation failed; unsaved chats were retained."))
      .finally(() => { busy = false; });
  });

  // One file holds everything this app can still reach, so it works even when saving is failing.
  on("history-export", async () => {
    await store.flush().catch(() => {});
    const saved = await getClient().historyExport().catch(() => null);
    const pending = store.pendingExport();
    const legacy = store.legacyUnmigrated();
    if (!saved && !pending.length && legacy === null) throw new Error("History could not be read. Reconnect and try again.");
    download("office-partner-history.json", { sessions: saved?.sessions ?? [], ...(pending.length ? { unsaved: pending } : {}), ...(legacy === null ? {} : { oldBrowserHistory: legacy }) });
    say(saved ? "History exported." : "Saved history is unavailable. Exported the unsaved and old browser copies only.");
  });

  on("history-recover-pending", async () => {
    if (isBusy()) throw new Error("Wait for the current response to finish.");
    const records = await store.recoverPending();
    recover(records);
    say(`Recovered ${records.length} unsaved chats as separate copies. Saved originals were kept.`);
  });

  return { refresh };
}
