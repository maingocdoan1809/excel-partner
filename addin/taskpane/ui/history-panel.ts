import type { ProviderId } from "../../../protocol/src/index.js";
import type { BridgeClient } from "../bridge-client.js";
import { formatWhen, type SessionStore, type StoredMessage, type StoredSession } from "./sessions.js";

export interface HistoryPanelDeps {
  getClient(): BridgeClient;
  getSession(): StoredSession;
  getRequestId(): string | null;
  store: SessionStore;
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

  function renderHistory(refreshTitles = true): void {
    deps.persistNow(); // so the current chat is listed too
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
        open.dataset.current = String(item.id === deps.getSession().id);
        const title = document.createElement("span");
        title.className = "history-title";
        title.textContent = item.title;
        const meta = document.createElement("span");
        meta.className = "history-meta";
        meta.textContent = [
          item.id === deps.getSession().id ? "Current" : null,
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
  }

  async function refreshHistoryTitles(): Promise<void> {
    if (!deps.getClient()?.connected) return;
    try {
      const summaries = await deps.getClient().usageSummaries();
      let changed = false;
      for (const summary of summaries) {
        if (!summary.title) continue;
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

  function openSession(id: string): void {
    deps.persistNow();
    const stored = store.get(id);
    if (!stored) return renderHistory();
    viewing = stored;
    els.sessionTitle.textContent = stored.title;
    els.sessionMeta.textContent = `${PROVIDER_NAMES[stored.provider as ProviderId] ?? stored.provider} · ${new Date(stored.startedAt).toLocaleString()}`;
    els.sessionTranscript.replaceChildren();
    for (const message of stored.messages) deps.renderMessage(els.sessionTranscript, message);
    els.sessionContinue.textContent = stored.id === deps.getSession().id ? "Back to chat" : "Continue this chat";
    deps.showPanel("session");
    els.sessionTranscript.scrollTop = 0;
  }

  function continueViewedSession(): void {
    if (!viewing) return;
    if (viewing.id !== deps.getSession().id) {
      const requestId = deps.getRequestId();
      if (requestId) deps.getClient().cancel(requestId);
      deps.persistNow();
      deps.loadSession(viewing);
    }
    viewing = null;
    deps.showPanel(null);
    els.input.focus();
  }

  function deleteViewedSession(): void {
    if (viewing) deleteSession(viewing.id);
  }

  function deleteSession(id: string): void {
    store.delete(id);
    if (id === deps.getSession().id) {
      deps.abandonRequest();
      // The open chat was removed: carry on in an empty one.
      deps.resetChat();
      void deps.refreshUsage();
    }
    viewing = null;
    deps.showPanel("history");
  }

  let confirmClearTimer: number | undefined;
  function clearHistory(): void {
    // Two clicks instead of a dialog (dialogs are unreliable inside task panes).
    if (els.historyClear.dataset.confirm !== "1") {
      els.historyClear.dataset.confirm = "1";
      els.historyClear.textContent = "Click again to delete all";
      confirmClearTimer = window.setTimeout(resetClearButton, 3_000);
      return;
    }
    window.clearTimeout(confirmClearTimer);
    resetClearButton();
    store.clear();
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
