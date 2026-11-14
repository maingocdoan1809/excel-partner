import { DurableHistory } from "../ui/durable-history.js";
import { activeHost } from "../host/index.js";
import { BrowserView } from "../ui/browser-view.js";
import { byId } from "../ui/dom.js";
import { showToast } from "../ui/toast.js";
import { FollowEnd, steadyVisibility } from "../ui/follow-end.js";
import { setupThinking } from "../ui/thinking.js";
import { ThoughtView } from "../ui/thought.js";
import { setupArtifactPreview } from "../ui/artifact-preview.js";
import { createHistoryScheduler } from "../ui/history-scheduler.js";
import { createSelfUpdater } from "../ui/self-update.js";
import { setupProviderLogin } from "../ui/provider-login.js";
import { setupProviderInstall } from "../ui/provider-install.js";
import { createSlashMenu } from "../ui/slash-menu.js";
import { setupConnectedFiles } from "../ui/connected-files.js";
import { setupSavedReferences } from "../ui/saved-references.js";
import { setupUserSkills } from "../ui/user-skills.js";
import { setupTestingPanel } from "../ui/testing-panel.js";
import { setupUpdatePanel } from "../ui/update-panel.js";
import { reloadKeepingChat } from "../ui/chat-restore.js";
import { setupMcp } from "../ui/mcp.js";
import { setupProviderOptions } from "../ui/provider-options.js";
import { setupSettings } from "../ui/settings.js";
import { setupHistoryStorage } from "../ui/history-storage.js";
import { setupSystemPanel } from "../ui/system-panel.js";
import { setupPanels } from "../ui/panels.js";
import { PROVIDER_NAMES, setupProviderPicker } from "../ui/provider-picker.js";
import { setupToolSettings } from "../ui/tool-settings.js";
import { setupTemplatePanel } from "../ui/template-panel.js";
import { setupUsagePanel } from "../ui/usage-panel.js";
import { setupChangesPanel } from "../ui/changes-panel.js";
import { setupHistoryPanel } from "../ui/history-panel.js";
import { newSession } from "../ui/sessions.js";
import { setupWorkspaceMenu } from "../ui/workspace-menu.js";
import { setupCellNavigation } from "../ui/cell-navigation.js";
import { els } from "../ui/elements.js";
import { append, renderMessage, renderWelcome, scrollToEnd } from "./messages.js";
import { loadSession, persistNow, schedulePersist, startNewChat } from "./chats.js";
import { abandonRequest } from "./turn.js";
import { currentWorkbookName, readWorkbookName } from "./workbook.js";
import { renderComposerPreview } from "./composer.js";
import { app, pendingAttachments } from "./state.js";

export const historyNotice = (text: string) => {
  const status = document.getElementById("history-save-status");
  if (status) { status.textContent = text; status.hidden = text === "History saved on this computer."; }
  const settingsStatus = document.getElementById("history-storage-status");
  if (settingsStatus) settingsStatus.textContent = text;
};
export const store = new DurableHistory(() => app.client, () => activeHost().id, localStorage, historyNotice);
export const browserView = new BrowserView({
  dock: byId("browser-dock"), dockLabel: byId("browser-dock-label"), dockStep: byId("browser-dock-step"), dockUrl: byId("browser-dock-url"),
  dockImage: byId<HTMLImageElement>("browser-dock-image"), dockOpen: byId<HTMLButtonElement>("browser-dock-open"),
  dockExpand: byId<HTMLButtonElement>("browser-dock-expand"), dockClose: byId<HTMLButtonElement>("browser-dock-close"),
  panelLabel: byId("browser-panel-label"), panelStep: byId("browser-panel-step"), panelUrl: byId("browser-panel-url"),
  panelImage: byId<HTMLImageElement>("browser-panel-image"), panelPrev: byId<HTMLButtonElement>("browser-prev"), panelNext: byId<HTMLButtonElement>("browser-next"),
}, () => showPanel("browser"), (count) => {
  // Closing the small view must not lose the pictures: the menu keeps a way back while there are any.
  byId("menu-browser").hidden = count === 0;
  byId("browser-count").textContent = String(count);
});
// Back to the chat with the small view above the message box, as it first appeared; its expand button opens the large one.
byId("menu-browser").addEventListener("click", () => { showPanel(null); browserView.showDock(); });
// While the person reads older messages the chat stays put; this button takes them back to the newest one.
export const jumpToLatest = document.createElement("button");
jumpToLatest.type = "button";
jumpToLatest.className = "jump-latest";
jumpToLatest.hidden = true;
jumpToLatest.textContent = "↓ Latest";
jumpToLatest.setAttribute("aria-label", "Jump to the latest message");
els.messages.after(jumpToLatest);
export const follow = new FollowEnd(els.messages, { onChange: steadyVisibility((visible) => { jumpToLatest.hidden = !visible; }) });
jumpToLatest.addEventListener("click", () => follow.scroll(true));
export const thinking = setupThinking(els.messages, { scrollToEnd: () => scrollToEnd(), onClear: () => thoughts.reset() });
export const thoughts = new ThoughtView(els.messages, () => thinking.element, () => scrollToEnd());
export const preview = setupArtifactPreview(() => showPanel("preview"));
export const historyScheduler = createHistoryScheduler(() => {
  app.session.updatedAt = Date.now();
  store.save(app.session);
  try { sessionStorage.setItem("excelPartner.activeChatId", app.session.id); } catch {}
});
export const updater = createSelfUpdater({ isBusy: () => app.currentRequestId !== null });
export const loginPanel = setupProviderLogin({
  client: () => app.client,
  provider: () => app.provider,
  loginKind: () => providerPicker.capabilitiesOf(app.provider)?.loginKind,
  canLogout: () => providerPicker.capabilitiesOf(app.provider)?.logout ?? false,
  isBusy: () => app.currentRequestId !== null,
  onSuccess: (signedIn) => {
    providerPicker.signedIn(signedIn);
    if (app.provider !== signedIn) selectProvider(signedIn);
    else {
      app.client.setProvider(signedIn);
      providerPicker.invalidateCatalog(signedIn);
      void refreshModelCatalog(signedIn);
    }
    showPanel(null);
  },
});
export const loginInstall = setupProviderInstall({
  client: () => app.client,
  provider: () => app.provider,
  isBusy: () => app.currentRequestId !== null,
  onInstalled: (installed, needsSignIn) => {
    if (installed !== app.provider) return;
    app.client.setProvider(installed);
    providerPicker.invalidateCatalog(installed);
    if (needsSignIn) providerPicker.requireLogin(installed);
    void refreshModelCatalog(installed);
    showPanel(needsSignIn ? "login" : null);
  },
});
export const slash = createSlashMenu({
  input: els.input,
  menu: els.slashMenu,
  onBuiltin: (name) => name === "clear" && startNewChat(),
  onSelect: (name) => { app.selectedSlashToken = `/${name}`; },
  getProvider: () => app.provider,
  loadCommands: () => app.client.listCommands(app.provider),
  onUnavailable: reason => append({ kind: "note", text: reason }),
});
export const connectedFiles = setupConnectedFiles({ onOpen: () => app.client?.refreshConnectedFiles() });
export const savedReferences = setupSavedReferences(() => app.client);
export const userSkills = setupUserSkills(() => app.client, () => slash.reload());
export const testingPanel = setupTestingPanel(() => app.client);
export const updatePanel = setupUpdatePanel({ getClient: () => app.client, reload: () => reloadKeepingChat(), appName: () => activeHost().app, announce: (text) => append({ kind: "note", text }) });
export const mcpSettings = setupMcp(() => app.client);
/** The settings the provider's CLI offers for the open session, drawn from what it reports. */
export const providerOptions = setupProviderOptions({
  container: byId("provider-options"),
  change: (id, value) => app.client.setProviderOption(id, value),
  isBusy: () => app.currentRequestId !== null,
  fail: message => showToast(message),
});
export const settings = setupSettings({
  loadInstructions: () => app.client.systemPrompt(),
  saveInstructions: (content, revision) => app.client.saveSystemPrompt(content, revision),
  historyCount: () => store.total || store.list().length,
  clearHistory: () => els.historyClear.click(), historyClearButton: els.historyClear,
});
export const historyStorage = setupHistoryStorage(() => app.client, store, records => { const recovered = records.find(r => r.originalId === app.session.id); if (recovered && !app.currentRequestId) loadSession(recovered.session); }, () => app.currentRequestId !== null);
export const systemPanel = setupSystemPanel(() => app.client, () => app.currentRequestId !== null);

export const { showPanel, openPanel, backFromPanel, openFloating, closeFloating, openMenu, closeMenu } = setupPanels({
  getProvider: () => app.provider, getProviderStatus: id => providerPicker.statusOf(id), getCapabilities: id => providerPicker.capabilitiesOf(id), elements: els, loginPanel, loginInstall,
  renderHistory: () => renderHistory(), renderChanges: () => renderChanges(), renderProviderMenu: () => renderProviderMenu(),
  openSettings: returning => {
    void historyStorage.refresh();
    void systemPanel.refresh();
    void mcpSettings.refresh();
    persistNow();
    void settings.open(returning);
    if (!returning) void updatePanel.refresh();
  },
  closePreview: () => preview.close(),
  refreshReferences: () => savedReferences.refresh(), refreshSkills: () => userSkills.refresh(),
  refreshTesting: () => testingPanel.refresh(), loadTemplateSkills: () => loadTemplateSkills(),
});

export const providerPicker = setupProviderPicker({
  getClient: () => app.client, getProvider: () => app.provider,
  setProvider: next => { app.provider = next; app.activeModel = undefined; providerOptions.clear(); },
  closeMenu, openPanel, startNewChat, refreshTools: () => toolSettings.refreshProvider(), append, elements: els,
});
export const toolSettings = setupToolSettings(() => providerPicker.capabilitiesOf(app.provider)?.nativeWebSearch ?? false);
export const { renderProviderMenu, renderAuthNotice, selectProvider, renderModelPicker, refreshModelCatalog, saveModelChoice, chosenModel } = providerPicker;

export const { loadTemplateSkills, renderTemplates, addTemplate } = setupTemplatePanel({ getClient: () => app.client, elements: els });

export const usagePanel = setupUsagePanel({
  getClient: () => app.client, getSession: () => app.session, getProvider: () => app.provider,
  showsCredits: () => providerPicker.capabilitiesOf(app.provider)?.credits ?? false,
  store, elements: els, persistNow, append,
});
export const { refreshUsage, renderUsage, exportUsage } = usagePanel;

export const { recordChange, renderChanges } = setupChangesPanel({
  getSession: () => app.session, schedulePersist, goTo: target => navigation.goTo(target), elements: els,
});

export const { renderHistory, continueViewedSession, deleteViewedSession, clearHistory } = setupHistoryPanel({
  getClient: () => app.client, getSession: () => app.session, getRequestId: () => app.currentRequestId,
  store, providerNames: PROVIDER_NAMES, elements: els,
  persistNow, abandonRequest, refreshUsage, loadSession, renderMessage, showPanel,
  resetChat: () => {
    app.session = newSession(app.provider);
    app.needsContext = false;
    els.messages.replaceChildren();
    renderWelcome();
    renderChanges();
  },
});

export const { updateMentionMenu, ensureWorkspaces, resetWorkspaces, chooseLocalFiles, setupFileDrop } = setupWorkspaceMenu({
  getClient: () => app.client, getWorkbookName: currentWorkbookName, closeSlash: () => slash.close(),
  renderComposerPreview, append, attachments: pendingAttachments, elements: els,
});
export const navigation = setupCellNavigation({ readName: readWorkbookName, goToFile: target => app.client.goToFile(target) });
