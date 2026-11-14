import { newSession, type StoredSession } from "../ui/sessions.js";
import { ProviderIdSchema } from "../../../protocol/src/index.js";
import { PROVIDER_KEY } from "../ui/provider-picker.js";
import { browserView, historyScheduler, providerOptions, refreshModelCatalog, refreshUsage, renderAuthNotice, renderChanges, renderModelPicker, renderProviderMenu, showPanel, slash, thinking } from "./services.js";
import { abandonRequest } from "./turn.js";
import { clearAttachment } from "./composer.js";
import { app, liveTools } from "./state.js";
import { els } from "../ui/elements.js";
import { renderMessage, renderWelcome, scrollToEnd } from "./messages.js";

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
  app.live = null;
  app.liveToolGroup = null;
  thinking.clear();
  liveTools.clear();
  els.messages.replaceChildren();
  renderWelcome();
  renderChanges();
  void refreshUsage();
  if (!keepPanel) showPanel(null);
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
  app.live = null;
  app.liveToolGroup = null;
  thinking.clear();
  liveTools.clear();
  els.messages.replaceChildren();
  for (const message of stored.messages) renderMessage(els.messages, message);
  scrollToEnd(true);
  renderChanges();
  void refreshUsage();
}
