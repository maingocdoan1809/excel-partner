import type { ChatHistoryStore } from "./durable-history.js";
import type { ProviderId } from "../../../protocol/src/index.js";
import type { BridgeClient } from "../bridge-client.js";
import { formatWhen, type StoredMessage, type StoredSession } from "./sessions.js";

export interface HistoryPanelDeps {
  getClient(): BridgeClient;
  getSession(): StoredSession;
  getRequestId(): string | null;
  store: ChatHistoryStore;
  providerNames: Record<ProviderId, string>;
  persistNow(): void;
  abandonRequest(): void;
  resetChat(): void;
  refreshUsage(): Promise<void>;
  loadSession(session: StoredSession): void;
  renderMessage(container: HTMLElement, message: StoredMessage): HTMLElement;
  showPanel(name: "session" | "history" | null): void;
  elements: {
    historyClear: HTMLButtonElement;
    historyList: HTMLElement;
    sessionTitle: HTMLElement;
    sessionMeta: HTMLElement;
    sessionTranscript: HTMLElement;
    sessionContinue: HTMLButtonElement;
    input: HTMLTextAreaElement;
  };
}

/** Saved chats and the transcript of the chat being viewed. The active chat stays owned by main. */
export function setupHistoryPanel(deps: HistoryPanelDeps) {
  const els = deps.elements;
  const store = deps.store;
  const PROVIDER_NAMES = deps.providerNames;
  let viewing: StoredSession | null = null;
  let transcriptOffset = 0, transcriptRevision = 0;
  const currentId = () => store.savedId?.(deps.getSession().id) ?? deps.getSession().id;

  async function renderHistory(refreshTitles = true, more = false): Promise<void> {
    deps.persistNow(); // so the current chat is listed too
    let pendingWarning = "";
    try { await store.flush?.(); } catch { pendingWarning = "Some chats are not saved yet. They are kept and saved again automatically; use Settings to export or recover them."; }
    try { await store.refresh?.(more); }
    catch (error) { els.historyList.textContent = `History is unavailable: ${error instanceof Error ? error.message : String(error)}. Trying again automatically. Unsaved chats are retained.`; return; }
    if (refreshTitles) void refreshHistoryTitles();
    const items = store.list();
    els.historyClear.hidden = items.length === 0;
    if (items.length === 0) {
      els.historyList.innerHTML = '<div class="empty">No saved chats yet. Your conversations are kept on this computer.</div>';
      return;
    }
    els.historyList.replaceChildren(
      ...items.map((item) => {
        const row = document.createElement("div");
        row.style.position = "relative";
        const open = document.createElement("button");
        open.type = "button";
        open.className = "history-item";
        open.dataset.current = String(item.id === currentId());
        const title = document.createElement("span");
        title.className = "history-title";
        title.textContent = item.title;
        const meta = document.createElement("span");
        meta.className = "history-meta";
        meta.textContent = [
          item.id === currentId() ? "Current" : null,
          PROVIDER_NAMES[item.provider as ProviderId] ?? item.provider,
          item.providerSessionId ? null : "Legacy chat",
          formatWhen(item.updatedAt),
          `${item.messageCount} message${item.messageCount === 1 ? "" : "s"}`,
          item.changeCount > 0 ? `${item.changeCount} change${item.changeCount === 1 ? "" : "s"}` : null,
        ]
          .filter(Boolean)
          .join(" · ");
        open.append(title, meta);
        open.addEventListener("click", () => openSession(item.id));
        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "history-delete";
        remove.title = "Delete this chat";
        remove.setAttribute("aria-label", "Delete this chat");
        remove.textContent = "✕";
        remove.addEventListener("click", () => deleteSession(item.id));
        row.append(open, remove);
        return row;
      }),
    );
    if (pendingWarning) { const warning = document.createElement("p"); warning.className = "settings-description"; warning.textContent = pendingWarning; els.historyList.prepend(warning); }
    if ((store.total ?? items.length) > items.length) {
      const moreButton = document.createElement("button"); moreButton.type = "button"; moreButton.className = "secondary"; moreButton.textContent = "Load older chats";
      moreButton.addEventListener("click", () => void renderHistory(false, true)); els.historyList.append(moreButton);
    }
  }

  async function refreshHistoryTitles(): Promise<void> {
    if (!deps.getClient()?.connected) return;
    try {
      const summaries = await deps.getClient().usageSummaries();
      let changed = false;
      for (const summary of summaries) {
        if (!summary.title || !store.list().some(s => s.id === summary.sessionId)) continue;
        const stored = summary.sessionId === deps.getSession().id ? deps.getSession() : store.get(summary.sessionId);
        if (!stored || stored.providerTitle === summary.title) continue;
        stored.providerTitle = summary.title;
        stored.title = summary.title;
        store.save(stored);
        changed = true;
      }
      if (changed) renderHistory(false);
    } catch { /* The local history remains usable while disconnected. */ }
  }

  async function openSession(id: string): Promise<void> {
    try {
    deps.persistNow();
    try { await store.flush?.(); } catch { /* Pending copies remain available from Settings; saved history can still be read. */ }
    const first = store.loadPage ? await store.loadPage(id, 0, 100) : null;
    transcriptOffset = first ? Math.max(0, first.totalMessages - 100) : 0;
    const page = first && transcriptOffset ? await store.loadPage!(id, transcriptOffset, 100) : first;
    if (first && page && first.revision !== page.revision) throw new Error("History changed in another window. Open it again.");
    const stored = page?.session ?? store.get(id);
    if (!stored) return renderHistory();
    transcriptRevision = page?.revision ?? 0;
    viewing = stored;
    els.sessionTitle.textContent = stored.title;
    els.sessionMeta.textContent = `${PROVIDER_NAMES[stored.provider as ProviderId] ?? stored.provider} · ${new Date(stored.startedAt).toLocaleString()}`;
    renderTranscript();
    els.sessionContinue.textContent = stored.id === currentId() ? "Back to chat" : "Continue this chat";
    deps.showPanel("session");
    els.sessionTranscript.scrollTop = 0;
    } catch (error) { els.historyList.textContent = `Could not open history: ${error instanceof Error ? error.message : String(error)}`; }
  }

  function renderTranscript() {
    els.sessionTranscript.replaceChildren();
    if (!viewing) return;
    if (transcriptOffset > 0 && store.loadPage) {
      const older = document.createElement("button"); older.type = "button"; older.className = "secondary"; older.textContent = "Load earlier messages";
      older.addEventListener("click", () => {
        const current = viewing; older.disabled = true;
        void (async () => {
          const offset = Math.max(0, transcriptOffset - 100);
          const page = await store.loadPage!(current!.id, offset, transcriptOffset - offset);
          if (viewing !== current) return;
          if (page.revision !== transcriptRevision) throw new Error("History changed. Open this chat again.");
          viewing!.messages.unshift(...page.session.messages); transcriptOffset = offset; renderTranscript();
        })().catch(error => { older.disabled = false; older.textContent = `Could not load messages: ${error.message}`; });
      });
      els.sessionTranscript.append(older);
    }
    for (const message of viewing.messages) deps.renderMessage(els.sessionTranscript, message);
  }

  async function continueViewedSession(): Promise<void> {
    if (!viewing) return;
    if (viewing.id !== currentId()) {
      const requestId = deps.getRequestId();
      if (requestId) deps.getClient().cancel(requestId);
      deps.persistNow();
      try {
        const complete = store.load ? await store.load(viewing.id) : viewing;
        if (!complete) throw new Error("Chat is unavailable.");
        deps.loadSession(complete);
      } catch (error) { els.sessionMeta.textContent = `Could not continue chat: ${error instanceof Error ? error.message : String(error)}`; return; }
    }
    viewing = null;
    deps.showPanel(null);
    els.input.focus();
  }

  function deleteViewedSession(): void {
    if (viewing) deleteSession(viewing.id);
  }

  async function deleteSession(id: string): Promise<void> {
    const deletingCurrent = id === currentId();
    try { await store.delete(id); } catch (error) { els.historyList.textContent = `Could not delete chat: ${error instanceof Error ? error.message : String(error)}`; return; }
    if (deletingCurrent) {
      deps.abandonRequest();
      // The open chat was removed: carry on in an empty one.
      deps.resetChat();
      void deps.refreshUsage();
    }
    viewing = null;
    deps.showPanel("history");
    void renderHistory();
  }

  let confirmClearTimer: number | undefined;
  async function clearHistory(): Promise<void> {
    // Two clicks instead of a dialog (dialogs are unreliable inside task panes).
    if (els.historyClear.dataset.confirm !== "1") {
      els.historyClear.dataset.confirm = "1";
      els.historyClear.textContent = "Click again to delete all";
      confirmClearTimer = window.setTimeout(resetClearButton, 3_000);
      return;
    }
    window.clearTimeout(confirmClearTimer);
    resetClearButton();
    try { await store.clear(); } catch (error) { els.historyList.textContent = `Could not clear history: ${error instanceof Error ? error.message : String(error)}`; return; }
    deps.abandonRequest();
    deps.resetChat();
    renderHistory();
    void deps.refreshUsage();
  }

  function resetClearButton(): void {
    els.historyClear.dataset.confirm = "0";
    els.historyClear.textContent = "Delete all";
  }

  return { renderHistory, continueViewedSession, deleteViewedSession, clearHistory };
}
