import { newSession, type StoredSession } from "../ui/sessions.js";
import { ProviderIdSchema } from "../../../protocol/src/index.js";
import { PROVIDER_KEY } from "../ui/provider-picker.js";
import { currentDocumentUrl } from "../ui/document-url.js";
import { browserView, historyScheduler, providerOptions, refreshModelCatalog, refreshUsage, renderAuthNotice, renderChanges, renderModelPicker, renderProviderMenu, showPanel, slash } from "./services.js";
import { abandonRequest } from "./turn.js";
import { clearAttachment } from "./composer.js";
import { prepareRuntime } from "./prepare.js";
import { app } from "./state.js";
import { els } from "../ui/elements.js";
import { renderMessage, renderWelcome, scrollToEnd } from "./messages.js";
import { store } from "./services.js";

export function schedulePersist(): void { historyScheduler.schedule(); }
export function persistNow(): void { historyScheduler.flush(); }
export function startNewChat(keepPanel = false): void {
  providerOptions.clear();
  slash.reload();
  abandonRequest();
  clearAttachment();
  persistNow();
  app.session = newSession(app.provider);
  browserView.reset();
  app.activeModel = undefined;
  app.needsContext = false;
  els.messages.replaceChildren();
  renderWelcome();
  renderChanges();
  void refreshUsage();
  if (!keepPanel) showPanel(null);
  prepareRuntime();
}

/** Make a stored chat the current one (its messages are shown in the chat area). */
export function loadSession(stored: StoredSession): void {
  providerOptions.clear();
  slash.reload();
  abandonRequest();
  clearAttachment();
  const storedProvider = ProviderIdSchema.safeParse(stored.provider).data;
  if (storedProvider && storedProvider !== app.provider) {
    app.provider = storedProvider;
    localStorage.setItem(PROVIDER_KEY, app.provider);
    if (app.client?.connected) app.client.setProvider(app.provider);
    renderProviderMenu();
    renderAuthNotice();
    renderModelPicker();
    void refreshModelCatalog(app.provider);
  }
  app.session = stored;
  try { sessionStorage.setItem("excelPartner.activeChatId", app.session.id); } catch {}
  browserView.reset();
  app.activeModel = undefined;
  app.needsContext = true;
  els.messages.replaceChildren();
  for (const message of stored.messages) renderMessage(els.messages, message);
  scrollToEnd(true);
  renderChanges();
  void refreshUsage();
  prepareRuntime();
}

let documentChatTried = false;

/** Once per pane: show the last chat held in the open file, if the pane is still empty. Another file's chats are never taken. */
export async function restoreDocumentChat(): Promise<void> {
  if (documentChatTried) return;
  documentChatTried = true;
  if (app.currentRequestId || app.session.messages.length) return;
  const documentKey = await currentDocumentUrl();
  const found = documentKey ? store.latestForDocument(documentKey) : null;
  if (!found) return;
  const stored = await store.loadActive(found.id);
  if (stored && !app.currentRequestId && !app.session.messages.length) loadSession(stored);
}
