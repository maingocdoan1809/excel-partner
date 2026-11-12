import { setupPanels } from "./ui/panels.js";
import { showToast } from "./ui/toast.js";
import { setupProviderPicker, loadProvider, PROVIDER_KEY, PROVIDER_NAMES } from "./ui/provider-picker.js";
import { setupWorkspaceMenu } from "./ui/workspace-menu.js";
import { byId } from "./ui/dom.js";
import { setupChangesPanel } from "./ui/changes-panel.js";
import { setupUsagePanel } from "./ui/usage-panel.js";
import { setupHistoryPanel } from "./ui/history-panel.js";
import { setupTemplatePanel } from "./ui/template-panel.js";
import { watchShapeSelectionEvents } from "./executor/selection-events.js";
import { activeHost, initHost } from "./host/index.js";
import { HOST_BASELINE, isSupported } from "./host/requirements.js";
import { monitorSelection } from "./ui/selection.js";
import type { ProviderEvent, ProviderId } from "../../protocol/src/index.js";
import { BridgeClient, randomId, type UploadedAttachment } from "./bridge-client.js";
import type { PermissionMode } from "./permissions.js";
import { UndoJournal } from "./executor/journal.js";
import { setProposalReviewer } from "./executor/word/index.js";
import { showProposalReview } from "./ui/proposal.js";
import { setupUsagePopover } from "./ui/usage-popover.js";
import { setupInputResizer } from "./ui/input-resizer.js";
import { renderMarkdown } from "./ui/markdown.js";
import { buildPreviewDocument, renderArtifactCard, type StoredArtifact } from "./ui/artifact.js";
import { BrowserView } from "./ui/browser-view.js";
import { setupTestingPanel } from "./ui/testing-panel.js";
import { setupUpdatePanel } from "./ui/update-panel.js";
import { reloadKeepingChat, takeChatRestore } from "./ui/chat-restore.js";
import { createSelfUpdater } from "./ui/self-update.js";
import { setupProviderLogin } from "./ui/provider-login.js";
import { setupProviderInstall } from "./ui/provider-install.js";
import { setupMcp } from "./ui/mcp.js";
import { setupToolSettings } from "./ui/tool-settings.js";
import { setupComposerMenu } from "./ui/composer-menu.js";
import { setupTheme } from "./ui/theme.js";
import { setupSavedReferences } from "./ui/saved-references.js";
import { setupUserSkills } from "./ui/user-skills.js";
import { showUserInput } from "./ui/user-input.js";
import { createHistoryScheduler } from "./ui/history-scheduler.js";
import { DurableHistory } from "./ui/durable-history.js";
import { setupHistoryStorage } from "./ui/history-storage.js";
import { setupSettings } from "./ui/settings.js";
import { setupSystemPanel } from "./ui/system-panel.js";
import { renderToolCard, paintTool, formatToolValue, toolChevron } from "./ui/tool-view.js";
import { ThoughtView } from "./ui/thought.js";
import { mentionsFile, mentionsSheet } from "./ui/cell-links.js";
import { ProviderIdSchema, exposedToolNamesForHost, sameWorkbookName } from "../../protocol/src/index.js";
import {
  buildContextPreface,
  bindProviderSession,
  newSession,
  type StoredMessage,
  type StoredSession,
  type StoredTool,
} from "./ui/sessions.js";
import { createSlashMenu } from "./ui/slash-menu.js";
import { leadingCommand, slashBadgeRange } from "./ui/slash.js";
import { parseNativeCommand, type ProviderCommand } from "../../protocol/src/commands.js";
import { atomicBadgeDeletion, fileMentionRanges } from "./ui/file-mentions.js";
import { filesFromClipboard } from "./ui/transfer-files.js";
import { setupConnectedFiles } from "./ui/connected-files.js";
import { delegatedPrompt } from "./ui/delegation.js";
import { renderReceivedMessage } from "./ui/received-message.js";

const els = {
  usageStats: byId<HTMLElement>("usage-stats"),
  usageBreakdown: byId<HTMLElement>("usage-breakdown"),
  usageCredits: byId<HTMLElement>("usage-credits"),
  usageCacheSplit: byId<HTMLElement>("usage-cache-split"),
  usageSource: byId<HTMLElement>("usage-source"),
  usageExport: byId<HTMLButtonElement>("usage-export"),
  status: byId<HTMLSpanElement>("connection-status"),
  statusLabel: byId<HTMLSpanElement>("connection-label"),
  messages: byId<HTMLElement>("messages"),
  menuButton: byId<HTMLButtonElement>("menu-button"),
  menuDot: byId<HTMLElement>("menu-dot"),
  menu: byId<HTMLElement>("settings-menu"),
  providerHint: byId<HTMLElement>("provider-hint"),
  menuChanges: byId<HTMLButtonElement>("menu-changes"),
  menuHistory: byId<HTMLButtonElement>("menu-history"),
  menuReferences: byId<HTMLButtonElement>("menu-references"),
  menuTemplates: byId<HTMLButtonElement>("menu-templates"),
  menuNew: byId<HTMLButtonElement>("menu-new"),
  changesCount: byId<HTMLSpanElement>("changes-count"),
  panels: {
    settings: byId<HTMLElement>("panel-settings"),
    skills: byId<HTMLElement>("panel-skills"),
    login: byId<HTMLElement>("panel-login"),
    changes: byId<HTMLElement>("panel-changes"),
    history: byId<HTMLElement>("panel-history"),
    references: byId<HTMLElement>("panel-references"),
    templates: byId<HTMLElement>("panel-templates"),
    session: byId<HTMLElement>("panel-session"),
    preview: byId<HTMLElement>("panel-preview"),
    browser: byId<HTMLElement>("panel-browser"),
    testing: byId<HTMLElement>("panel-testing"),
  },
  changesList: byId<HTMLElement>("changes-list"),
  historyList: byId<HTMLElement>("history-list"),
  historyClear: byId<HTMLButtonElement>("history-clear"),
  templateSkill: byId<HTMLSelectElement>("template-skill"),
  templateAdd: byId<HTMLButtonElement>("template-add"),
  templateInput: byId<HTMLInputElement>("template-input"),
  templateList: byId<HTMLElement>("template-list"),
  sessionTitle: byId<HTMLElement>("session-title"),
  sessionMeta: byId<HTMLElement>("session-meta"),
  sessionTranscript: byId<HTMLElement>("session-transcript"),
  sessionBack: byId<HTMLButtonElement>("session-back"),
  sessionDelete: byId<HTMLButtonElement>("session-delete"),
  sessionContinue: byId<HTMLButtonElement>("session-continue"),
  resizer: byId<HTMLElement>("composer-resizer"),
  chip: byId<HTMLDivElement>("selection-chip"),
  input: byId<HTMLTextAreaElement>("prompt-input"),
  modelSelect: byId<HTMLSelectElement>("model-select"),
  effortSelect: byId<HTMLSelectElement>("effort-select"),
  permissionMode: byId<HTMLSelectElement>("permission-mode"),
  slashMenu: byId<HTMLDivElement>("slash-menu"),
  mentionMenu: byId<HTMLDivElement>("mention-menu"),
  promptDisplay: byId<HTMLDivElement>("prompt-display"),
  attachButton: byId<HTMLButtonElement>("attach-button"),
  send: byId<HTMLButtonElement>("send-button"),
  sendLabel: document.querySelector<HTMLSpanElement>("#send-button .label")!,
};

const PERMISSION_MODE_KEY = "excelPartner.permissionMode";
const NARROW_BELOW = 330;
/** After the pane reloaded itself, a chat younger than this is reopened automatically. */
const RESTORE_WITHIN_MS = 12 * 3_600_000;

const historyNotice = (text: string) => {
  const status = document.getElementById("history-save-status");
  if (status) { status.textContent = text; status.hidden = text === "History saved on this computer."; }
  const settingsStatus = document.getElementById("history-storage-status");
  if (settingsStatus) settingsStatus.textContent = text;
};
const store = new DurableHistory(() => client, () => activeHost().id, localStorage, historyNotice);
let restoreChatId: string | null = null;
const journal = new UndoJournal((tool) => activeHost().writeTools.has(tool), async () => client.saveBackup((await currentWorkbookName()) || activeHost().noun, await activeHost().readFile()));
// Handle for end-to-end checks against a live workbook; it grants nothing execute_office_js does not already have.
(globalThis as { __undoJournal?: UndoJournal }).__undoJournal = journal;
let client: BridgeClient;
let provider: ProviderId = loadProvider();
let activeModel: string | null | undefined;
let activeEffort: string | undefined;
let permissionMode: PermissionMode = localStorage.getItem(PERMISSION_MODE_KEY) === "auto" ? "auto" : "ask";
let session: StoredSession = newSession(provider);
let requestMessageIndex = 0;
let requestChangeIndex = 0;
let requestStartedAt = 0;
/** The provider process starts empty after a reload or when a past chat is continued; replay the transcript once. */
let needsContext = false;
let currentRequestId: string | null = null;
let requestSession: StoredSession | null = null;
let cancelRequested = false;
let promptDispatched = false;
interface PendingPrompt {
  id: string;
  parentRequestId: string;
  prompt: string;
  attachments: typeof pendingAttachments;
  message: StoredMessage;
  state: "pending" | "next_turn" | "held";
}
const pendingPrompts = new Map<string, PendingPrompt>();
let autoQueueAllowed = true;
let live: { element: HTMLElement; message: StoredMessage } | null = null;
const liveTools = new Map<string, { card: HTMLElement; tool: StoredTool }>();
// Tools the bridge answers by itself reach the pane only as provider events, never as a tool request, so they must keep
// their own card: the person sees which other file was read or sent a request.
const bridgeAnsweredTools: ReadonlySet<string> = new Set(["list_connected_files", "read_connected_file", "send_to_connected_file"]);
let workbookToolNamesCache: Set<string> | null = null;
/** Tools of the open file's application that reach the pane as tool requests (their card comes from the executed result). */
const isDocumentTool = (name: string): boolean => (workbookToolNamesCache ??= new Set<string>(exposedToolNamesForHost(activeHost().id, true, { playwright: true, webSearch: true }).filter((tool) => !bridgeAnsweredTools.has(tool)))).has(name);
const browserView = new BrowserView({
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
const previewFrame = byId<HTMLIFrameElement>("preview-frame");
const previewSource = byId<HTMLElement>("preview-source");
const previewToggle = byId<HTMLButtonElement>("preview-toggle");
const previewCopy = byId<HTMLButtonElement>("preview-copy");
let previewing: StoredArtifact | null = null;
let liveToolGroup: { element: HTMLDetailsElement; message: StoredMessage } | null = null;
let thinking: HTMLElement | null = null;
const thoughts = new ThoughtView(els.messages, () => thinking, () => scrollToEnd());
/** The other workbooks open with this add-in (not this one), as the bridge last listed them. */
let otherFiles: string[] = [];
/** Set once the assistant has used the connected-file tools in the running request: what it says may be about another file. */
let turnTouchedOtherFiles = false;
const CONNECTED_FILE_TOOLS: ReadonlySet<string> = new Set(["list_connected_files", "read_connected_file", "send_to_connected_file"]);
/** How to link the cells in a reply. Bare cells (no link written by the assistant) are linked only when it cannot be about another file. */
/** The names of this workbook's sheets, as last read from Excel (so a reply that names one is not guessed at). */
let sheetNames: string[] = [];
const cellLinkOptions = (text: string, aboutOtherFiles = false) => ({ plain: !aboutOtherFiles && !mentionsFile(text, otherFiles) && !mentionsSheet(text, sheetNames) });
const pendingAttachments: Array<{ reference: UploadedAttachment; name: string }> = [];
const historyScheduler = createHistoryScheduler(() => {
  session.updatedAt = Date.now();
  store.save(session);
  try { sessionStorage.setItem("excelPartner.activeChatId", session.id); } catch {}
});
let permissionPicker: ReturnType<typeof setupComposerMenu> | undefined;
let selectedSlashToken: string | null = null;

const updater = createSelfUpdater({ isBusy: () => currentRequestId !== null });
const loginPanel = setupProviderLogin({
  client: () => client,
  provider: () => provider,
  isBusy: () => currentRequestId !== null,
  onSuccess: (signedIn) => {
    providerPicker.signedIn(signedIn);
    if (provider !== signedIn) selectProvider(signedIn);
    else {
      client.setProvider(signedIn);
      providerPicker.invalidateCatalog(signedIn);
      void refreshModelCatalog(signedIn);
    }
    showPanel(null);
  },
});
const loginInstall = setupProviderInstall({
  client: () => client,
  provider: () => provider,
  isBusy: () => currentRequestId !== null,
  onInstalled: (installed, needsSignIn) => {
    if (installed !== provider) return;
    client.setProvider(installed);
    providerPicker.invalidateCatalog(installed);
    if (needsSignIn) providerPicker.requireLogin(installed);
    void refreshModelCatalog(installed);
    showPanel(needsSignIn ? "login" : null);
  },
});
const slash = createSlashMenu({
  input: els.input,
  menu: els.slashMenu,
  onBuiltin: (name) => name === "clear" && startNewChat(),
  onSelect: (name) => { selectedSlashToken = `/${name}`; },
  getProvider: () => provider,
  loadCommands: () => client.listCommands(provider),
  onUnavailable: reason => append({ kind: "note", text: reason }),
});
const connectedFiles = setupConnectedFiles({ onOpen: () => client?.refreshConnectedFiles() });
const savedReferences = setupSavedReferences(() => client);
const userSkills = setupUserSkills(() => client, () => slash.reload());
const testingPanel = setupTestingPanel(() => client);
const updatePanel = setupUpdatePanel({ getClient: () => client, reload: () => reloadKeepingChat(), appName: () => activeHost().app, announce: (text) => append({ kind: "note", text }) });
const mcpSettings = setupMcp(() => client);
const toolSettings = setupToolSettings(() => provider);
const settings = setupSettings({
  loadInstructions: () => client.systemPrompt(),
  saveInstructions: (content, revision) => client.saveSystemPrompt(content, revision),
  historyCount: () => store.total || store.list().length,
  clearHistory: () => els.historyClear.click(), historyClearButton: els.historyClear,
});
const historyStorage = setupHistoryStorage(() => client, store, records => { const recovered = records.find(r => r.originalId === session.id); if (recovered && !currentRequestId) loadSession(recovered.session); }, () => currentRequestId !== null);
const systemPanel = setupSystemPanel(() => client, () => currentRequestId !== null);

const { showPanel, openPanel, backFromPanel, openFloating, closeFloating, openMenu, closeMenu } = setupPanels({
  getProvider: () => provider, getProviderStatus: id => providerPicker.statusOf(id), elements: els, loginPanel, loginInstall,
  renderHistory: () => renderHistory(), renderChanges: () => renderChanges(), renderProviderMenu: () => renderProviderMenu(),
  openSettings: returning => {
    void historyStorage.refresh();
    void systemPanel.refresh();
    void mcpSettings.refresh();
    persistNow();
    void settings.open(returning);
    if (!returning) void updatePanel.refresh();
  },
  closePreview: () => { previewFrame.srcdoc = ""; previewing = null; },
  refreshReferences: () => savedReferences.refresh(), refreshSkills: () => userSkills.refresh(),
  refreshTesting: () => testingPanel.refresh(), loadTemplateSkills: () => loadTemplateSkills(),
});

const providerPicker = setupProviderPicker({
  getClient: () => client, getProvider: () => provider,
  setProvider: next => { provider = next; activeModel = undefined; },
  closeMenu, openPanel, startNewChat, refreshTools: () => toolSettings.refreshProvider(), append, elements: els,
});
const { renderProviderMenu, renderAuthNotice, selectProvider, renderModelPicker, refreshModelCatalog, saveModelChoice } = providerPicker;

const { loadTemplateSkills, renderTemplates, addTemplate } = setupTemplatePanel({ getClient: () => client, elements: els });

const usagePanel = setupUsagePanel({
  getClient: () => client, getSession: () => session, getProvider: () => provider,
  store, elements: els, persistNow, append,
});
const { refreshUsage, renderUsage, exportUsage } = usagePanel;

const { recordChange, renderChanges } = setupChangesPanel({
  getSession: () => session, journal, schedulePersist, goTo, elements: els,
});

const { renderHistory, continueViewedSession, deleteViewedSession, clearHistory } = setupHistoryPanel({
  getClient: () => client, getSession: () => session, getRequestId: () => currentRequestId,
  store, providerNames: PROVIDER_NAMES, elements: els,
  persistNow, abandonRequest, refreshUsage, loadSession, renderMessage, showPanel,
  resetChat: () => {
    session = newSession(provider);
    needsContext = false;
    els.messages.replaceChildren();
    renderWelcome();
    renderChanges();
  },
});

const { updateMentionMenu, ensureWorkspaces, resetWorkspaces, chooseLocalFiles, setupFileDrop } = setupWorkspaceMenu({
  getClient: () => client, getWorkbookName: currentWorkbookName, closeSlash: () => slash.close(),
  renderComposerPreview, append, attachments: pendingAttachments, elements: els,
});

Office.onReady((info) => {
  // Excel, Word and PowerPoint share this pane. Any other Office application has no tools here.
  if (!["Excel", "Word", "PowerPoint"].includes(String(info.host))) return;
  initHost(String(info.host));
  applyHostBranding();
  {
    // The manifest only asks for what every host has; what this host needs is checked here, so an old Office says so.
    const [set, version] = HOST_BASELINE[activeHost().id];
    if (!isSupported([set, version])) window.setTimeout(() => showToast(`Office Partner needs ${set} ${version} in ${activeHost().app}, which this version of Office does not have. Update Office to use it.`), 0);
  }

  client = new BridgeClient({
    onUserInput: async (request, signal, requestId) => {
      if (requestId !== currentRequestId || signal.aborted) return { cancelled: true, answers: {} };
      clearThinking();
      live = null;
      const questionText = request.questions.map(q => q.question).join("\n\n");
      const questionBubble = append({ kind: "assistant", text: questionText });
      const owner = session;
      const dock = byId("question-dock");
      dock.hidden = false;
      document.body.classList.add("question-active");
      const pending = showUserInput(dock, request, signal);
      scrollToEnd();
      const response = await pending;
      dock.hidden = true;
      document.body.classList.remove("question-active");
      if (session === owner) questionBubble.innerHTML = renderMarkdown(questionText, { cellLinks: cellLinkOptions(questionText) });
      if (session === owner && requestId === currentRequestId) {
        append({ kind: response.cancelled ? "note" : "user", text: response.cancelled ? "Question skipped or cancelled." : request.questions.map(q =>
          `${q.header || q.question}: ${q.isSecret ? "[Hidden answer]" : response.answers[q.id]!.join(", ")}`).join("\n") });
        if (!signal.aborted) showThinking();
      }
      return response;
    },
    onState: renderState,
    onHello: ({ addinBuildId }) => {
      void store.connect().then(async () => {
        if (restoreChatId && !currentRequestId) {
          const id = restoreChatId; restoreChatId = null;
          const restored = await store.loadActive(id); if (restored && !currentRequestId) loadSession(restored);
        }
      }).catch(error => historyNotice(`History is unavailable: ${error.message}. Trying again automatically.`));
      void mcpSettings.restore();
      void publishWorkbookName();
      void refreshSheetNames();
      slash.reload();
      client.setProvider(provider);
      void refreshUsage();
      void refreshModelCatalog(provider);
      if (session.messages.length) needsContext = true;
      activeModel = undefined;
      resetWorkspaces();
      void ensureWorkspaces();
      updater.check(addinBuildId);
      void updatePanel.afterReload();
    },
    onProviderStatus: ({ provider: id, detection }) => providerPicker.onStatus(id, detection),
    onProviderEvent: handleEvent,
    runTool: (context, tool, args, execute) => journal.run(context, tool, args, execute),
    onBrowserFrame: (frame) => { if (frame.requestId === currentRequestId) browserView.push(frame); },
    readWorkbookName,
    onConnectedFiles: (files) => {
      // Only workbooks count here: this list decides which cell references in a reply can be linked to another file.
      const names = files.filter((file) => !file.current && file.host === "excel").map((file) => file.name).filter(Boolean);
      const changed = names.join("\n") !== otherFiles.join("\n");
      otherFiles = names;
      connectedFiles.render(files);
      // Whether a reply's bare cells are linked depends on which other files are open (a chat opened before this list arrived, for one).
      if (changed && !currentRequestId && session.messages.length) redrawMessages();
    },
    // Another file's chat asked this workbook to show a cell (a link there): only the selection moves.
    goTo: selectRange,
    onDelegatedPrompt: (request) => {
      // The other assistant's request is a turn like any other here, so it waits its turn: refuse while this chat is busy.
      if (currentRequestId || cancelRequested || resolvingCommand) {
        client.declineDelegation(request.delegationId, "That file is busy with its own request.");
        return;
      }
      void submit(undefined, { delegationId: request.delegationId, from: request.from.name, fromHost: request.from.host ?? "excel", text: request.text });
    },
    onError: (message) => showToast(message),
    onToolExecuted: (tool, args, result, context) => {
      if (tool === "create_artifact") {
        // The card is the visible result; a second raw tool entry would only repeat the file contents.
        const file = args as { filename?: unknown; content?: unknown };
        if (result.ok && context.requestId === currentRequestId && session === requestSession && typeof file.filename === "string" && typeof file.content === "string") {
          liveToolGroup = null;
          append({ kind: "artifact", text: "", artifact: { name: file.filename, content: file.content } });
          scrollToEnd();
        }
        return;
      }
      if (context.requestId === currentRequestId && session === requestSession) appendTurnTool({ name: tool, args: summarize(args), input: formatToolValue(args),
        result: formatToolValue(result), state: result.ok ? "ok" : "error", error: result.ok ? null : result.error.message });
      recordChange(tool, args, result, context);
      if (result.ok && ["add_shape", "add_connector", "add_line", "update_shape", "delete_shapes", "execute_office_js", "manage_sheet", "insert_sheets_from_file"].includes(tool)) {
        refreshSelectionSubscriptions?.();
      }
    },
  });
  client.setPermissionMode(permissionMode);
  client.connect();

  setupTheme(byId<HTMLSelectElement>("theme-select"));
  providerPicker.setupModelPicker();
  const savedEffort = localStorage.getItem("excelPartner.effort");
  els.effortSelect.value = savedEffort === "low" || savedEffort === "high" ? savedEffort : "medium";
  const effortButton = byId<HTMLButtonElement>("effort-button");
  const refreshEffort = () => {
    const effort = els.effortSelect.value;
    const label = els.effortSelect.selectedOptions[0]!.textContent;
    effortButton.dataset.effort = effort;
    effortButton.setAttribute("aria-label", `Reasoning effort: ${label}`);
    els.effortSelect.title = `Reasoning effort: ${label}. Models without effort control use provider defaults. Changes apply to your next message.`;
    effortButton.title = els.effortSelect.title;
    localStorage.setItem("excelPartner.effort", effort);
  };
  refreshEffort();
  els.effortSelect.addEventListener("change", refreshEffort);
  setupComposerMenu(els.effortSelect, effortButton, byId("effort-menu"));
  permissionPicker = setupComposerMenu(els.permissionMode, byId<HTMLButtonElement>("permission-button"), byId("permission-menu"));

  watchSelection();
  updater.poll();
  setupInputResizer(els.resizer, els.input);
  const observer = new ResizeObserver(([entry]) => {
    document.body.dataset.narrow = String((entry?.contentRect.width ?? window.innerWidth) < NARROW_BELOW);
  });
  observer.observe(document.body);

  // Settings menu
  els.menuButton.addEventListener("click", (event) => {
    event.stopPropagation();
    showPanel("settings");
  });
  document.addEventListener("click", (event) => {
    for (const popover of document.querySelectorAll<HTMLElement>('[data-floating-open="true"]')) {
      if (!popover.contains(event.target as Node) && !document.getElementById(popover.dataset.floatingAnchor ?? "")?.contains(event.target as Node)) closeFloating(popover);
    }
    if (!els.menu.hidden && !els.menu.contains(event.target as Node)) closeMenu();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    const fallback = document.querySelector<HTMLElement>('[data-floating-open="true"]');
    if (fallback) { closeFloating(fallback); document.getElementById(fallback.dataset.floatingAnchor ?? "")?.focus(); return; }
    if (!els.menu.hidden) closeMenu();
    else if (!slash.isOpen && openPanel() !== null) backFromPanel();
  });
  byId<HTMLSelectElement>("provider-select").addEventListener("change", event => selectProvider((event.target as HTMLSelectElement).value as ProviderId));
  els.menuChanges.addEventListener("click", () => showPanel("changes"));
  // Cell and range mentions in the assistant's replies (see ui/cell-links.ts).
  els.messages.addEventListener("click", (event) => {
    const link = (event.target as HTMLElement | null)?.closest<HTMLAnchorElement>("a.cell-link");
    if (!link?.dataset.address) return;
    event.preventDefault();
    void goToLink({
      ...(link.dataset.file ? { file: link.dataset.file } : {}),
      ...(link.dataset.sheet ? { sheet: link.dataset.sheet } : {}),
      address: link.dataset.address,
    });
  });
  byId("menu-login").addEventListener("click", () => showPanel("login"));
  byId("provider-login-button").addEventListener("click", () => showPanel("login"));
  byId("login-recheck").addEventListener("click", () => {
    client.setProvider(provider);
    void refreshModelCatalog(provider);
    showPanel(null);
  });
  byId("new-chat-button").addEventListener("click", () => startNewChat());
  byId("tab-chat").addEventListener("click", () => showPanel(null));
  byId("tab-changes").addEventListener("click", () => showPanel("changes"));
  byId("tab-history").addEventListener("click", () => showPanel("history"));
  els.menuHistory.addEventListener("click", () => showPanel("history"));
  byId("history-button").addEventListener("click", () => showPanel("history"));
  byId("menu-settings").addEventListener("click", () => showPanel("settings"));
  document.querySelectorAll<HTMLButtonElement>("[data-settings-panel]").forEach(button => button.addEventListener("click", () => {
    const panel = button.dataset.settingsPanel;
    if (panel === "changes" || panel === "references" || panel === "skills" || panel === "templates") showPanel(panel);
  }));
  els.menuReferences.addEventListener("click", () => showPanel("references"));
  els.menuTemplates.addEventListener("click", () => showPanel("templates"));
  byId("menu-skills").addEventListener("click", () => showPanel("skills"));
  byId("menu-testing").addEventListener("click", () => showPanel("testing"));
  byId("changes-button").addEventListener("click", () => showPanel("changes"));
  byId("tools-button").addEventListener("click", () => openFloating(byId("composer-settings-menu"), byId("tools-button")));
  byId("composer-skills").addEventListener("click", () => {
    void userSkills.refresh();
    openFloating(els.panels.skills, byId("tools-button"));
  });
  byId("composer-tools").addEventListener("click", () => {
    void mcpSettings.refresh();
    const popover = byId("composer-tools-popover");
    popover.append(byId("tools-menu"));
    openFloating(popover, byId("tools-button"));
  });
  els.menuNew.addEventListener("click", () => {
    closeMenu();
    startNewChat();
  });

  // Panels
  document.querySelectorAll<HTMLButtonElement>("[data-close]").forEach((button) => button.addEventListener("click", backFromPanel));
  els.sessionBack.addEventListener("click", () => showPanel("history"));
  els.sessionContinue.addEventListener("click", continueViewedSession);
  els.sessionDelete.addEventListener("click", deleteViewedSession);
  els.historyClear.addEventListener("click", clearHistory);
  els.templateSkill.addEventListener("change", () => void renderTemplates());
  els.templateAdd.addEventListener("click", () => els.templateInput.click());
  els.templateInput.addEventListener("change", () => {
    const file = els.templateInput.files?.[0];
    els.templateInput.value = "";
    if (file) void addTemplate(file);
  });

  els.usageExport.addEventListener("click", () => void exportUsage());
  setupUsagePopover(byId<HTMLButtonElement>("usage-button"), byId("usage-popover"), () => void refreshUsage());
  // Composer
  els.mentionMenu.addEventListener("mousedown", (event) => {
    if (!(event.target instanceof HTMLInputElement)) event.preventDefault();
  });
  // Keep the caret in the message box: if the button takes focus, the box's blur handler closes the @ menu.
  els.attachButton.addEventListener("mousedown", (event) => event.preventDefault());
  els.attachButton.addEventListener("click", () => {
    if (els.mentionMenu.hidden) {
      const cursor = els.input.selectionStart;
      const before = els.input.value.slice(0, cursor);
      const inserted = `${before}${before && !/\s$/.test(before) ? " " : ""}@`;
      els.input.value = `${inserted}${els.input.value.slice(cursor)}`;
      els.input.focus();
      els.input.setSelectionRange(inserted.length, inserted.length);
      renderComposerPreview();
      updateMentionMenu();
    } else els.mentionMenu.hidden = true;
  });
  els.permissionMode.value = permissionMode;
  permissionPicker.refresh();
  els.permissionMode.addEventListener("change", () => selectPermissionMode(els.permissionMode.value as PermissionMode));
  els.modelSelect.addEventListener("change", saveModelChoice);
  els.send.addEventListener("click", () => {
    if (currentRequestId && !els.input.value.trim() && !pendingAttachments.length) stopCurrentTurn();
    else void submit();
  });
  byId("stop-button").addEventListener("click", stopCurrentTurn);
  els.input.addEventListener("keydown", (event) => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === "Enter" && event.ctrlKey) {
      event.preventDefault();
      els.input.setRangeText("\n", els.input.selectionStart, els.input.selectionEnd, "end");
      els.input.dispatchEvent(new Event("input", { bubbles: true }));
      return;
    }
    if (event.key === "Escape" && !els.mentionMenu.hidden) {
      event.preventDefault();
      els.mentionMenu.hidden = true;
      return;
    }
    if (slash.handleKey(event)) return;
    if (event.key === "Enter" && !event.shiftKey && !event.altKey && !event.metaKey) {
      event.preventDefault();
      if (!event.repeat) void submit();
    }
  });
  els.input.addEventListener("beforeinput", (event) => {
    if (!event.inputType.startsWith("delete")) return;
    const text = els.input.value;
    const badges: Array<{ start: number; end: number }> = fileMentionRanges(text, pendingAttachments.map((item) => item.name));
    const command = slashBadgeRange(text, selectedSlashToken);
    if (command) badges.push(command);
    const range = atomicBadgeDeletion(
      text,
      els.input.selectionStart,
      els.input.selectionEnd,
      event.inputType,
      badges,
    );
    if (!range) return;
    event.preventDefault();
    els.input.setRangeText("", range.start, range.end, "start");
    syncAttachmentsWithInput();
    syncSelectedSlashToken();
    renderComposerPreview();
    slash.update();
    updateMentionMenu();
  });
  els.input.addEventListener("input", () => {
    syncAttachmentsWithInput();
    syncSelectedSlashToken();
    renderComposerPreview();
    slash.update();
    updateMentionMenu();
  });
  els.input.addEventListener("scroll", () => { els.promptDisplay.scrollTop = els.input.scrollTop; });
  els.input.addEventListener("paste", (event) => {
    const files = filesFromClipboard(event.clipboardData);
    if (!files.length) return;
    event.preventDefault();
    void chooseLocalFiles(files);
  });
  setupFileDrop();
  els.input.addEventListener("blur", () => window.setTimeout(() => {
    slash.close();
    // The box may already have its focus back (for example after the + button), so only close when focus left both.
    const focused = document.activeElement;
    if (focused !== els.input && !els.mentionMenu.contains(focused)) els.mentionMenu.hidden = true;
  }, 120));
  window.addEventListener("pagehide", () => { persistNow(); void store.flush(true).catch(() => {}); });
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") { persistNow(); void store.flush().catch(() => {}); } });

  // Bring the last chat back only after the pane reloaded itself (a new build or a finished update). A pane that was
  // closed and opened again starts a new chat; the old one is in the history list.
  const restoring = takeChatRestore();
  try { restoreChatId = restoring ? sessionStorage.getItem("excelPartner.activeChatId") : null; } catch {}
  const previous = restoring ? (restoreChatId ? store.getActive(restoreChatId) : store.latest(RESTORE_WITHIN_MS)) : null;
  if (previous) { restoreChatId ??= previous.id; loadSession(previous); }
  else renderWelcome();
  renderProviderMenu();
  renderModelPicker();
  renderChanges();
  void refreshUsage();
});

/* -------------------------------------------------------------------------- */
/* Messages: one renderer for live chat, restored chat and the history viewer   */
/* -------------------------------------------------------------------------- */

function renderMessage(container: HTMLElement, message: StoredMessage): HTMLElement {
  // Older saved chats may contain tool-only assistant records with no text.
  // Keep them out of the transcript as well as newly streamed empty messages.
  if (message.kind === "assistant" && !message.text.trim()) return document.createElement("div");
  container.querySelector(".welcome")?.remove();
  let element: HTMLElement;
  switch (message.kind) {
    case "user": {
      if (message.from !== undefined) { element = renderReceivedMessage(message.from, message.text, message.fromHost); break; }
      element = bubble("user");
      const lead = leadingCommand(message.text);
      if (lead) {
        const token = document.createElement("span");
        token.className = "slash-token";
        token.textContent = lead.command;
        element.append(token);
        appendFileMentions(element, lead.rest, message.attachments?.map((item) => item.name) ?? []);
      } else appendFileMentions(element, message.text, message.attachments?.map((item) => item.name) ?? []);
      if (message.delivery) renderDelivery(element, message);
      break;
    }
    case "assistant":
      element = bubble("assistant");
      element.innerHTML = renderMarkdown(message.text, { cellLinks: cellLinkOptions(message.text, message.aboutOtherFiles) });
      break;
    case "error":
      element = bubble("error");
      element.textContent = message.text;
      break;
    case "note":
      element = document.createElement("div");
      element.className = "note";
      element.textContent = message.text;
      break;
    case "tool":
      element = renderToolCard(message.tool!);
      break;
    case "artifact":
      element = renderArtifactCard(message.artifact!, openArtifact);
      break;
    case "tool_group": {
      const group = document.createElement("details");
      group.className = "tool-group";
      const summary = document.createElement("summary");
      const label = document.createElement("span");
      label.textContent = toolGroupLabel(message.tools?.length ?? 0);
      const chevron = document.createElement("span");
      chevron.className = "tool-group-chevron";
      chevron.append(toolChevron());
      summary.append(label, chevron);
      const list = document.createElement("div");
      list.className = "tool-list";
      for (const tool of message.tools ?? []) list.append(renderToolCard(tool));
      group.append(summary, list);
      element = group;
      break;
    }
  }
  if (container === els.messages && thinking) container.insertBefore(element, thinking);
  else container.append(element);
  return element;
}


function toolGroupLabel(count: number): string {
  return `${count} tool${count === 1 ? "" : "s"} called`;
}

/** One disclosure per provider turn, shared by native and workbook tool calls. */
function appendTurnTool(tool: StoredTool): HTMLElement {
  if (!liveToolGroup) {
    const message: StoredMessage = { kind: "tool_group", text: "", tools: [] };
    liveToolGroup = { element: append(message) as HTMLDetailsElement, message };
  }
  liveToolGroup.message.tools!.push(tool);
  liveToolGroup.element.querySelector("summary > span")!.textContent = toolGroupLabel(liveToolGroup.message.tools!.length);
  const card = renderToolCard(tool);
  liveToolGroup.element.querySelector(".tool-list")!.append(card);
  schedulePersist();
  return card;
}

function bubble(kind: "user" | "assistant" | "error"): HTMLDivElement {
  const element = document.createElement("div");
  element.className = `bubble ${kind}`;
  return element;
}


/** Add a message to the current chat, show it and save. */
function append(message: StoredMessage): HTMLElement {
  session.messages.push(message);
  const element = renderMessage(els.messages, message);
  scrollToEnd();
  schedulePersist();
  return element;
}

function renderWelcome(): void {
  const box = document.createElement("div");
  box.className = "welcome";
  const illustration = document.createElement("img");
  illustration.className = "welcome-sheet";
  illustration.src = "../assets/spreadsheet-chalk.svg";
  illustration.alt = "Hand-drawn spreadsheet";
  illustration.draggable = false;
  box.append(illustration);
  els.messages.append(box);
}

/* -------------------------------------------------------------------------- */
/* Conversation                                                                */
/* -------------------------------------------------------------------------- */

function renderDelivery(element: HTMLElement, message: StoredMessage): void {
  element.querySelector(".message-delivery")?.remove();
  if (!message.delivery) return;
  const label = document.createElement("span");
  label.className = "message-delivery";
  const pending = message.queueId ? pendingPrompts.get(message.queueId) : undefined;
  label.textContent = message.delivery === "pending" && pending?.state !== "held" && pending ? "Pending" : "Not sent";
  if (!pending || pending.state === "held") {
    const send = document.createElement("button");
    send.type = "button";
    send.textContent = "Send";
    send.addEventListener("click", () => {
      if (currentRequestId || !client.connected) return;
      void submit(pending ?? { id: message.queueId ?? randomId(), parentRequestId: "", prompt: message.text, attachments: [], message, state: "held" });
    });
    label.append(send);
  }
  element.append(label);
  if (message.queueId) element.dataset.queueId = message.queueId;
}

function updateDelivery(item: PendingPrompt): void {
  item.message.delivery = item.state === "held" ? "held" : "pending";
  for (const element of els.messages.querySelectorAll<HTMLElement>(".bubble.user[data-queue-id]")) {
    if (element.dataset.queueId === item.id) renderDelivery(element, item.message);
  }
  schedulePersist();
}

function holdPendingPrompts(): void {
  for (const item of pendingPrompts.values()) { item.state = "held"; updateDelivery(item); }
}

function sendNextPending(): void {
  if (currentRequestId || !client.connected) return;
  const next = [...pendingPrompts.values()].find(item => item.state === "next_turn" && session.messages.includes(item.message));
  if (next) void submit(next);
}

let resolvingCommand = false;

async function runClientCommand(command: ProviderCommand, args: string): Promise<void> {
  switch (command.action) {
    case "new": startNewChat(); break;
    case "history": showPanel("history"); break;
    case "model":
      if (args) {
        await refreshModelCatalog(provider);
        if (!Array.from(els.modelSelect.options).some(option => option.value === args)) throw new Error(`Unknown model: ${args}`);
        els.modelSelect.value = args; saveModelChoice(); renderModelPicker();
      } else byId<HTMLButtonElement>("model-button").click();
      break;
    case "effort": byId<HTMLButtonElement>("effort-button").click(); break;
    case "permissions":
      if (args === "ask" || args === "auto") selectPermissionMode(args);
      else if (args) throw new Error("Use /permissions ask or /permissions auto.");
      else byId<HTMLButtonElement>("permission-button").click();
      break;
    case "usage": byId<HTMLButtonElement>("usage-button").click(); break;
    case "login": showPanel("login"); break;
    case "theme": {
      const select = byId<HTMLSelectElement>("theme-select");
      if (args) {
        if (!["light", "dark", "auto"].includes(args)) throw new Error("Use /theme light, dark, or auto.");
        select.value = args; select.dispatchEvent(new Event("change"));
      } else { openMenu(); select.focus(); }
      break;
    }
    case "copy": {
      const answer = [...session.messages].reverse().find(message => message.kind === "assistant");
      if (!answer) throw new Error("There is no assistant response to copy yet.");
      await navigator.clipboard.writeText(answer.text);
      append({ kind: "note", text: "Copied the last response." });
      break;
    }
    case "export": {
      const markdown = session.messages.filter(message => message.kind !== "tool_group").map(message => `## ${message.kind}\n\n${message.text}`).join("\n\n");
      const url = URL.createObjectURL(new Blob([markdown], { type: "text/markdown;charset=utf-8" }));
      const link = document.createElement("a"); link.href = url; link.download = "excel-partner-chat.md";
      document.body.append(link); link.click(); link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      break;
    }
    case "help": append({ kind: "note", text: slash.commands.map(item => `/${item.name} — ${item.description}${item.command?.execution === "unsupported" ? " (CLI only)" : ""}`).join("\n") }); break;
    case "status": append({ kind: "note", text: `Provider: ${PROVIDER_NAMES[provider]}\nModel: ${els.modelSelect.value || "Default"}\nSession: ${session.providerSessionId || "New chat"}\nPermissions: ${permissionMode}` }); break;
  }
}

/** `delegated`: another open workbook's assistant sent this request; it runs as a normal turn here and its reply goes back. */
async function submit(queued?: PendingPrompt, delegated?: { delegationId: string; from: string; fromHost?: "excel" | "word" | "powerpoint"; text: string }): Promise<void> {
  // Whoever sent a delegated request waits for an answer: every way out of here that does not start the turn must say so.
  const decline = (reason: string) => { if (delegated) client.declineDelegation(delegated.delegationId, reason); };
  const prompt = delegated ? delegatedPrompt(delegated.from, delegated.text)
    : queued?.prompt ?? (els.input.value.trim() || (pendingAttachments.length ? "Please inspect the attached files." : ""));
  if (!prompt || !client.connected) { decline("That file is not connected right now."); return; }
  if (prompt.length > 32_000) { decline("The request is too long for that file."); showToast("Messages must be 32,000 characters or fewer."); return; }
  let nativeCommand = false;
  if (parseNativeCommand(prompt)) {
    if (resolvingCommand) return;
    const owner = session, selectedProvider = provider;
    resolvingCommand = true;
    try {
      const item = await slash.resolve(prompt);
      if (session !== owner || provider !== selectedProvider || (!queued && els.input.value.trim() !== prompt)) return;
      if (!item) throw new Error("Unknown command. Open the / menu to see available commands and skills.");
      if (item.command?.execution === "unsupported") throw new Error(item.command.reason ?? "This command requires the CLI terminal.");
      if (item.command?.execution === "client" || item.kind === "builtin") {
        await runClientCommand(item.command ?? { name: item.name, description: item.description, source: "builtin", execution: "client", action: "new" }, parseNativeCommand(prompt)!.args);
        els.input.value = ""; slash.close(); renderComposerPreview(); return;
      }
      nativeCommand = item.kind === "provider";
      if (nativeCommand && currentRequestId) throw new Error("Wait for the current response to finish before running a provider command.");
      if (nativeCommand && pendingAttachments.length) throw new Error("Send attached files separately from a provider command.");
    } catch (error) { showToast(error instanceof Error ? error.message : String(error)); return; }
    finally { resolvingCommand = false; }
  }
  const turnAttachments = delegated ? [] as typeof pendingAttachments : queued?.attachments ?? pendingAttachments.slice();
  if (!delegated && !turnAttachments.length && /(?:^|\s)@[^\n]*\.(?:xlsx|xlsm|xltx|xltm|docx|docm|pptx|pptm|png|jpe?g|gif|webp|md|txt|json|csv|ya?ml|zip)(?=\s|$)/i.test(prompt)) {
    showToast("The @ file is only text. Choose it from a workspace or the local file picker so I can read it.");
    updateMentionMenu();
    return;
  }
  if (currentRequestId) {
    if (delegated) { decline("That file is busy with its own request."); return; }
    if (queued || cancelRequested) return;
    if (pendingPrompts.size >= 20) { showToast("Wait for pending messages to be sent before adding more."); return; }
    const id = randomId();
    const attachments = pendingAttachments.slice();
    const message: StoredMessage = { kind: "user", text: prompt, delivery: "pending", queueId: id, attachments: attachments.map(item => ({ name: item.name })) };
    const item: PendingPrompt = { id, parentRequestId: currentRequestId, prompt, attachments, message, state: "pending" };
    pendingPrompts.set(id, item);
    try {
      client.queuePrompt(currentRequestId, id, prompt,
        attachments.flatMap(item => item.reference.id ? [] : [item.reference.path]),
        attachments.flatMap(item => item.reference.id ? [item.reference.id] : []));
      append(message);
      els.input.value = "";
      clearAttachment();
      persistNow();
    } catch (error) { pendingPrompts.delete(id); showToast(error instanceof Error ? error.message : String(error)); }
    return;
  }
  if (!session.identityVersion && !session.providerSessionId && session.messages.some(message => message.kind === "user")) {
    decline("The chat open in that file is an older chat that cannot take requests. Ask the user to start a new chat in that window.");
    showToast("This legacy chat has no saved CLI session ID, so it cannot be resumed reliably. Its history is preserved. Start a new chat to create a native session.");
    return;
  }
  const model = els.modelSelect.value || null;
  const effort = els.effortSelect.value as "low" | "medium" | "high";
  saveModelChoice();
  slash.close();
  if (prompt === "/clear") {
    els.input.value = "";
    startNewChat();
    return;
  }

  const requestId = randomId();
  if (liveToolGroup) liveToolGroup.element.open = false;
  liveToolGroup = null;
  currentRequestId = requestId;
  requestSession = session;
  turnTouchedOtherFiles = false;
  browserView.reset();
  setBusy(true);
  if (!queued && !delegated) els.input.value = "";
  renderComposerPreview();
  showPanel(null);
  const context = !nativeCommand && (needsContext || (activeModel !== undefined && activeModel !== model) || (activeEffort !== undefined && activeEffort !== effort))
    ? buildContextPreface(session) : undefined;
  if (!nativeCommand) needsContext = false;
  requestMessageIndex = queued ? Math.max(0, session.messages.indexOf(queued.message)) : session.messages.length;
  requestChangeIndex = session.changes.length;
  requestStartedAt = Date.now();
  if (queued) {
    queued.message.delivery = undefined;
    pendingPrompts.delete(queued.id);
    for (const element of els.messages.querySelectorAll<HTMLElement>(".bubble.user[data-queue-id]")) {
      if (element.dataset.queueId === queued.id) renderDelivery(element, queued.message);
    }
  } else append(delegated ? { kind: "user", text: delegated.text, from: delegated.from, fromHost: delegated.fromHost }
    : { kind: "user", text: prompt, attachments: turnAttachments.map((item) => ({ name: item.name })) });
  if (!queued && !delegated) clearAttachment();
  showThinking();
  persistNow();

  const result = nativeCommand ? null : await activeHost().execute("get_selection", {});
  if (currentRequestId !== requestId || requestSession !== session) { decline("The chat open in that file changed before the request could start."); return; }
  const selection = result?.ok ? (result.data as never) : null;
  if (cancelRequested) return finishLocalCancel(requestId, delegated?.delegationId);
  try {
    await mcpSettings.ready();
    const attachmentIds = turnAttachments.flatMap((item) => item.reference.id ? [item.reference.id] : []);
    const referencePaths = turnAttachments.flatMap((item) => item.reference.id ? [] : [item.reference.path]);
    if (cancelRequested) return finishLocalCancel(requestId, delegated?.delegationId);
    client.sendPrompt(requestId, prompt, selection, context, attachmentIds, model, referencePaths, effort, session.id,
      session.provider === provider ? session.providerSessionId : undefined, toolSettings.value(), nativeCommand || undefined, delegated?.delegationId);
    promptDispatched = true;
    activeModel = model;
    activeEffort = effort;
  } catch (error) {
    if (cancelRequested) return finishLocalCancel(requestId, delegated?.delegationId);
    decline("The request could not be sent in that file.");
    clearThinking();
    currentRequestId = null;
    setBusy(false);
    if (queued) { queued.state = "held"; pendingPrompts.set(queued.id, queued); updateDelivery(queued); }
    else if (!delegated) els.input.value = prompt;
    renderComposerPreview();
    showToast(error instanceof Error ? error.message : String(error));
  }
}

function stopCurrentTurn(): void {
  if (!currentRequestId || cancelRequested) return;
  cancelRequested = true;
  autoQueueAllowed = false;
  els.send.dataset.stopping = "true";
  els.send.disabled = true;
  els.send.setAttribute("aria-label", "Stopping response");
  els.send.title = "Stopping response";
  if (promptDispatched) client.cancel(currentRequestId);
}

/** `delegationId`: the stopped turn answered another file's request, which is told it will not be done. */
function finishLocalCancel(requestId: string, delegationId?: string): void {
  if (currentRequestId !== requestId) return;
  if (delegationId) client.declineDelegation(delegationId, "The user stopped the request in that file before it started.");
  clearThinking();
  currentRequestId = null;
  holdPendingPrompts();
  const last = session.messages[session.messages.length - 1];
  if (last?.kind === "user" && pendingPrompts.size === 0) {
    session.messages.pop();
    if (last.from !== undefined) {
      // Another assistant's request is not the person's own text: drop its block and leave the composer alone.
      const received = els.messages.querySelectorAll(".received");
      received[received.length - 1]?.remove();
    } else {
      const bubbles = els.messages.querySelectorAll(".bubble.user");
      bubbles[bubbles.length - 1]?.remove();
      els.input.value = last.text;
      renderComposerPreview();
    }
  }
  needsContext = session.messages.length > 0;
  setBusy(false);
  append({ kind: "note", text: "Stopped." });
  persistNow();
  updater.reloadIfIdle();
}

async function currentWorkbookName(): Promise<string> {
  return (await activeHost().readName()) ?? "";
}

function appendFileMentions(container: HTMLElement, text: string, names: readonly string[]): void {
  let cursor = 0;
  for (const range of fileMentionRanges(text, names)) {
    container.append(document.createTextNode(text.slice(cursor, range.start)));
    const badge = document.createElement("span");
    badge.className = "file-mention";
    badge.textContent = text.slice(range.start, range.end);
    container.append(badge);
    cursor = range.end;
  }
  container.append(document.createTextNode(text.slice(cursor)));
}

function renderComposerPreview(): void {
  updateSendControl();
  els.promptDisplay.replaceChildren();
  const text = els.input.value;
  const command = slashBadgeRange(text, selectedSlashToken);
  if (command) {
    const badge = document.createElement("span");
    badge.className = "slash-token";
    badge.textContent = text.slice(command.start, command.end);
    els.promptDisplay.append(badge);
    appendFileMentions(els.promptDisplay, text.slice(command.end), pendingAttachments.map((item) => item.name));
  } else appendFileMentions(els.promptDisplay, text, pendingAttachments.map((item) => item.name));
  // A textarea keeps a caret line after a trailing newline; mirror that line too.
  if (text.endsWith("\n")) els.promptDisplay.append(document.createTextNode("\u200b"));
  els.promptDisplay.scrollTop = els.input.scrollTop;
}

function syncSelectedSlashToken(): void {
  if (selectedSlashToken && !(
    els.input.value === selectedSlashToken ||
    els.input.value.startsWith(`${selectedSlashToken} `) ||
    els.input.value.startsWith(`${selectedSlashToken}\n`)
  )) selectedSlashToken = null;
}

function syncAttachmentsWithInput(): void {
  const ranges = fileMentionRanges(els.input.value, pendingAttachments.map((item) => item.name));
  const counts = new Map<string, number>();
  for (const range of ranges) counts.set(range.name, (counts.get(range.name) ?? 0) + 1);
  for (let i = pendingAttachments.length - 1; i >= 0; i--) {
    const name = pendingAttachments[i]!.name;
    const left = counts.get(name) ?? 0;
    if (left > 0) counts.set(name, left - 1);
    else pendingAttachments.splice(i, 1);
  }
}

function clearAttachment(): void {
  pendingAttachments.length = 0;
  selectedSlashToken = null;
  renderComposerPreview();
}

function handleEvent(event: ProviderEvent): void {
  if (event.type !== "followup_status" && event.requestId !== null && event.requestId !== currentRequestId) return;
  switch (event.type) {
    case "followup_status": {
      const item = pendingPrompts.get(event.id);
      if (!item || item.parentRequestId !== event.requestId) return;
      if (item.state === "held" && event.state !== "delivered") return;
      if (event.state === "delivered") {
        item.message.delivery = undefined;
        pendingPrompts.delete(item.id);
        if (session.messages.includes(item.message)) {
          for (const element of els.messages.querySelectorAll<HTMLElement>(".bubble.user[data-queue-id]")) {
            if (element.dataset.queueId === item.id) renderDelivery(element, item.message);
          }
        }
        schedulePersist();
      } else {
        item.state = event.state === "rejected" ? "held" : event.state;
        if (!currentRequestId && !autoQueueAllowed) item.state = "held";
        updateDelivery(item);
        if (event.message) showToast(event.message);
        if (event.state === "next_turn" && autoQueueAllowed) sendNextPending();
      }
      return;
    }
    case "session_info": {
      if (event.requestId !== currentRequestId || session !== requestSession) return;
      const oldId = session.id;
      const bound = bindProviderSession(session, event, requestMessageIndex, requestChangeIndex, requestStartedAt);
      if (bound.previous) {
        store.save(bound.previous); // Removes a provisional empty history item, preserves older sessions.
        requestMessageIndex = 0;
        requestChangeIndex = 0;
      }
      session = bound.current;
      requestSession = session;
      if (oldId !== session.id) {
        usagePanel.rebind(event.requestId);
        // Metadata precedes output, so redraw only when a request moves to another native session.
        clearThinking();
        els.messages.replaceChildren();
        for (const message of session.messages) renderMessage(els.messages, message);
        showThinking();
        renderChanges();
      }
      persistNow();
      return;
    }
    case "session_starting":
      setThinkingLabel("Starting session…");
      return;
    case "provider_progress":
      if (event.requestId === currentRequestId) setThinkingLabel(event.message);
      return;
    case "turn_started":
      if (event.requestId !== currentRequestId) return;
      void refreshSheetNames();
      setThinkingLabel(null);
      usagePanel.start(event.requestId);
      return;
    case "usage_update":
      usagePanel.update(event.requestId, event.usage);
      return;
    case "thinking_delta":
      thoughts.add(event.text);
      return;
    case "message_delta":
      if (!event.text || (!live && !event.text.trim())) return;
      thoughts.end();
      if (!live) live = startLive("assistant", event.text);
      else live.message.text += event.text;
      live.element.innerHTML = renderMarkdown(live.message.text, { cellLinks: cellLinkOptions(live.message.text, live.message.aboutOtherFiles) });
      scrollToEnd();
      schedulePersist();
      return;
    case "message_completed":
      // Providers can finish an assistant message that only calls a tool.
      // It has no chat text and must not create (or erase) a visible bubble.
      if (!event.text.trim()) { live = null; return; }
      thoughts.end();
      if (!live) live = startLive("assistant", event.text);
      live.message.text = event.text;
      live.element.innerHTML = renderMarkdown(event.text, { cellLinks: cellLinkOptions(event.text, live.message.aboutOtherFiles) });
      {
        // The reply may name a sheet the assistant just made: read the names again and draw it once more if they changed.
        const { element, message } = live;
        void refreshSheetNames(false).then((changed) => {
          if (changed && element.isConnected) element.innerHTML = renderMarkdown(message.text, { cellLinks: cellLinkOptions(message.text, message.aboutOtherFiles) });
        });
      }
      live = null; // the next message (after a tool call) gets its own bubble
      scrollToEnd();
      schedulePersist();
      return;
    case "tool_started": {
      thoughts.end();
      live = null;
      if (CONNECTED_FILE_TOOLS.has(event.tool)) turnTouchedOtherFiles = true;
      // Workbook calls get their exact result from onToolExecuted. Provider-local
      // tools have a separate native call ID and must not be paired by name.
      if (isDocumentTool(event.tool)) return;
      const tool: StoredTool = { name: event.tool, args: summarize(event.args), input: formatToolValue(event.args), state: "running", error: null };
      const card = appendTurnTool(tool);
      liveTools.set(event.callId, { card, tool });
      schedulePersist();
      return;
    }
    case "tool_finished": {
      const entry = liveTools.get(event.callId);
      if (!entry) return;
      entry.tool.state = event.ok ? "ok" : "error";
      entry.tool.error = event.error;
      paintTool(entry.card, entry.tool);
      liveTools.delete(event.callId);
      schedulePersist();
      return;
    }
    case "turn_finished":
      if (event.requestId !== currentRequestId) return;
      autoQueueAllowed = event.reason === "completed";
      usagePanel.finish(event.requestId);
      void refreshUsage();
      slash.reload();
      void userSkills.refresh();
      clearThinking();
      currentRequestId = null;
      live = null;
      for (const { card, tool } of liveTools.values()) {
        tool.state = "error";
        tool.error = "no result";
        paintTool(card, tool);
      }
      liveTools.clear();
      if (liveToolGroup) liveToolGroup.element.open = false;
      liveToolGroup = null;
      setBusy(false);
      if (event.reason === "cancelled") append({ kind: "note", text: "Stopped." });
      if (event.reason !== "completed") holdPendingPrompts();
      persistNow();
      void store.flush().catch(() => {});
      sendNextPending();
      updater.reloadIfIdle();
      return;
    case "error":
      clearThinking();
      if (event.code === "login_required") {
        providerPicker.requireLogin(provider);
      }
      showToast(event.message);
      return;
    default:
      return;
  }
}

function showThinking(): void {
  if (thinking) return;
  thinking = document.createElement("div");
  thinking.className = "thinking";
  thinking.setAttribute("role", "status");
  thinking.setAttribute("aria-label", "AI is working");
  thinking.innerHTML = '<span class="thinking-dots" aria-hidden="true"><i></i><i></i><i></i></span>';
  els.messages.append(thinking);
  scrollToEnd();
}

/** Words next to the dots while something slow is happening (the provider CLI opening); null goes back to dots only. */
function setThinkingLabel(text: string | null): void {
  if (!thinking) return;
  thinking.querySelector(".thinking-label")?.remove();
  thinking.setAttribute("aria-label", text ?? "AI is working");
  if (text === null) return;
  const label = document.createElement("span");
  label.className = "thinking-label";
  label.textContent = text;
  thinking.append(label);
}

function clearThinking(): void {
  thoughts.reset();
  thinking?.remove();
  thinking = null;
}

function startLive(kind: "assistant", text: string): { element: HTMLElement; message: StoredMessage } {
  const message: StoredMessage = { kind, text, ...(turnTouchedOtherFiles ? { aboutOtherFiles: true } : {}) };
  session.messages.push(message);
  return { element: renderMessage(els.messages, message), message };
}

/* -------------------------------------------------------------------------- */
/* Chats: new, restore, persist                                                */
/* -------------------------------------------------------------------------- */

function schedulePersist(): void { historyScheduler.schedule(); }
function persistNow(): void { historyScheduler.flush(); }

function startNewChat(keepPanel = false): void {
  slash.reload();
  abandonRequest();
  clearAttachment();
  persistNow();
  session = newSession(provider);
  browserView.reset();
  activeModel = undefined;
  needsContext = false;
  live = null;
  liveToolGroup = null;
  clearThinking();
  liveTools.clear();
  els.messages.replaceChildren();
  renderWelcome();
  renderChanges();
  void refreshUsage();
  if (!keepPanel) showPanel(null);
}

/** Draw the current chat's messages again, where they were. Only between requests: a running one has live parts on screen. */
function redrawMessages(): void {
  const top = els.messages.scrollTop;
  els.messages.replaceChildren();
  for (const message of session.messages) renderMessage(els.messages, message);
  els.messages.scrollTop = top;
}

/** Make a stored chat the current one (its messages are shown in the chat area). */
function loadSession(stored: StoredSession): void {
  slash.reload();
  abandonRequest();
  clearAttachment();
  const storedProvider = ProviderIdSchema.safeParse(stored.provider).data;
  if (storedProvider && storedProvider !== provider) {
    provider = storedProvider;
    localStorage.setItem(PROVIDER_KEY, provider);
    if (client?.connected) client.setProvider(provider);
    renderProviderMenu();
    renderAuthNotice();
    renderModelPicker();
    void refreshModelCatalog(provider);
  }
  session = stored;
  try { sessionStorage.setItem("excelPartner.activeChatId", session.id); } catch {}
  browserView.reset();
  activeModel = undefined;
  needsContext = true;
  live = null;
  liveToolGroup = null;
  clearThinking();
  liveTools.clear();
  els.messages.replaceChildren();
  for (const message of stored.messages) renderMessage(els.messages, message);
  scrollToEnd();
  renderChanges();
  void refreshUsage();
}

/* -------------------------------------------------------------------------- */
/* Tool permissions                                                           */
/* -------------------------------------------------------------------------- */

function selectPermissionMode(mode: PermissionMode): void {
  if (mode !== "ask" && mode !== "auto") return;
  permissionMode = mode;
  localStorage.setItem(PERMISSION_MODE_KEY, mode);
  client.setPermissionMode(mode);
  els.permissionMode.value = mode;
  permissionPicker?.refresh();
}
/* -------------------------------------------------------------------------- */
/* Artifact preview                                                           */
/* -------------------------------------------------------------------------- */

function showSource(on: boolean): void {
  previewToggle.setAttribute("aria-pressed", String(on));
  previewToggle.textContent = on ? "Preview" : "Source";
  previewFrame.hidden = on;
  previewSource.hidden = !on;
}

function openArtifact(artifact: StoredArtifact): void {
  previewing = artifact;
  byId("preview-title").textContent = artifact.name;
  previewSource.textContent = artifact.content;
  previewFrame.srcdoc = buildPreviewDocument(artifact);
  showSource(false);
  showPanel("preview");
}

previewToggle.addEventListener("click", () => showSource(previewToggle.getAttribute("aria-pressed") !== "true"));
previewCopy.addEventListener("click", () => {
  if (!previewing) return;
  navigator.clipboard.writeText(previewing.content).then(
    () => { previewCopy.textContent = "Copied"; },
    () => { previewCopy.textContent = "Copy failed"; },
  ).finally(() => window.setTimeout(() => { previewCopy.textContent = "Copy"; }, 1500));
});

/**
 * Read the sheet names from Excel. True when they changed. With `redraw`, a changed list also draws the chat again, since which
 * cells of a reply are linked depends on it (a chat opened before the first read, for one).
 */
async function refreshSheetNames(redraw = true): Promise<boolean> {
  if (!activeHost().hasSheets) return false;
  try {
    const names = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map((sheet) => sheet.name);
    });
    if (names.join("\n") === sheetNames.join("\n")) return false;
    sheetNames = names;
    if (redraw && !currentRequestId && session.messages.length) redrawMessages();
    return true;
  } catch {
    return false; // no Excel behind this page, or it is busy: keep the names we have
  }
}

/** Show a range in this workbook: its sheet (or the open one) is activated and the range selected. Throws when it cannot. */
async function selectRange(target: { sheet?: string; address: string }): Promise<void> {
  if (!activeHost().hasSheets) throw new Error(`Links to cells only work in Excel; this is ${activeHost().app}.`);
  await Excel.run(async (context) => {
    const sheet = target.sheet ? context.workbook.worksheets.getItemOrNullObject(target.sheet) : context.workbook.worksheets.getActiveWorksheet();
    if (target.sheet) {
      await context.sync();
      if ((sheet as Excel.Worksheet & { isNullObject?: boolean }).isNullObject) throw new Error(`There is no sheet named "${target.sheet}" in this workbook.`);
    }
    sheet.activate();
    sheet.getRange(target.address).select();
    await context.sync();
  });
}

const placeLabel = (target: { file?: string; sheet?: string; address: string }) =>
  `${target.file ? `${target.file} → ` : ""}${target.sheet ? `${target.sheet}!` : ""}${target.address}`;

async function goTo(target: { sheet?: string; address: string }): Promise<void> {
  try {
    await selectRange(target);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    showToast(`Cannot go to ${target.sheet ? `${target.sheet}!` : ""}${target.address}: ${reason}`);
  }
}

/** A link in a reply: in this workbook, or, when it names another open file, in that file's window (the bridge asks its pane). */
async function goToLink(target: { file?: string; sheet?: string; address: string }): Promise<void> {
  if (!target.file) return goTo(target);
  const here = await readWorkbookName();
  if (here && sameWorkbookName(here, target.file)) return goTo({ ...(target.sheet ? { sheet: target.sheet } : {}), address: target.address });
  try {
    await client.goToFile({ file: target.file, ...(target.sheet ? { sheet: target.sheet } : {}), address: target.address });
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    showToast(`Cannot go to ${placeLabel(target)}: ${reason}`);
  }
}

/* -------------------------------------------------------------------------- */
/* Selection chip                                                              */
/* -------------------------------------------------------------------------- */

let refreshSelectionSubscriptions: (() => void) | undefined;

function watchSelection(): void {
  const monitor = monitorSelection({
    read: () => activeHost().readSelectionLabel(),
    isVisible: () => document.visibilityState !== "hidden",
    render: value => {
      els.chip.textContent = value.label;
      els.chip.setAttribute("aria-label", `Current selection: ${value.label}`);
      els.chip.title = value.title;
    },
    onError: error => {
      els.chip.textContent = (error as { code?: string })?.code === "InvalidSelection" ? "Object selected (shape / chart)" : "No selection";
      els.chip.setAttribute("aria-label", `Current selection: ${els.chip.textContent}`);
      els.chip.title = "Current selection could not be identified. Select an object again or focus this pane to retry.";
    },
  });
  // Let Excel finish its native activation/deactivation transition before reading
  // the active object. This is a one-shot debounce, not a polling timer.
  let selectionTimer: number | undefined;
  const selectionChanged = () => {
    window.clearTimeout(selectionTimer);
    selectionTimer = window.setTimeout(() => monitor.refresh(), 80);
  };
  // Shape selection events exist only for Excel; Word and PowerPoint rely on DocumentSelectionChanged, focus and the send-time read.
  const events = activeHost().hasSheets ? watchShapeSelectionEvents({
    onActivated: value => {
      window.clearTimeout(selectionTimer);
      monitor.show(value);
    },
    onSelection: selectionChanged,
    onError: error => { console.warn("Shape selection event registration failed", error); },
  }) : { refresh: async () => undefined, stop: async () => undefined };
  const refresh = () => { monitor.refresh(); void events.refresh(); };
  refreshSelectionSubscriptions = refresh;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectionChanged);
  window.addEventListener("focus", refresh);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState !== "hidden") refresh();
  });
  window.addEventListener("pagehide", () => {
    window.clearTimeout(selectionTimer);
    monitor.stop();
    refreshSelectionSubscriptions = undefined;
    void events.stop();
  }, { once: true });
}

/* -------------------------------------------------------------------------- */
/* Which Office application this is                                            */
/* -------------------------------------------------------------------------- */

/** Names and wording that depend on the host: the brand, the page title and the review card Word edits are approved in. */
function applyHostBranding(): void {
  const host = activeHost();
  document.title = "Office Partner";
  const brand = document.querySelector(".brand-text strong");
  if (brand) brand.textContent = "Office Partner";
  document.body.dataset.host = host.id;
  // Static wording in index.html names Excel; say the application this pane is really in.
  for (const el of document.querySelectorAll<HTMLElement>(".host-app")) el.textContent = host.app;
  for (const el of document.querySelectorAll<HTMLElement>("[data-host-label]")) el.textContent = (el.dataset.hostLabel ?? "").replace("{app}", host.app);
  if (!host.hasSheets) {
    // The testing tools work from a workbook only. Connected files work from every Office app.
    byId("menu-testing").hidden = true;
  }
  if (host.id === "word") {
    // propose_doc_edits stages edits; the card below is where the person applies them.
    setProposalReviewer((request, signal) => {
      const dock = byId("question-dock");
      dock.hidden = false;
      document.body.classList.add("question-active");
      const outcome = showProposalReview(dock, request, signal ?? new AbortController().signal);
      const close = () => { dock.hidden = true; document.body.classList.remove("question-active"); };
      void outcome.then(close, close);
      return outcome;
    });
  }
}

/* -------------------------------------------------------------------------- */
/* Workbook name, so other open workbooks can list this one                    */
/* -------------------------------------------------------------------------- */

/** What Office calls the open file now (it changes on Save As), or null when it cannot be read. There is no event for a rename. */
async function readWorkbookName(): Promise<string | null> {
  return activeHost().readName();
}

/**
 * Tell the bridge the name once on connect, so the file is listed with its name at once. After that the bridge asks
 * (`readWorkbookName`) when someone looks at the list, so nothing here runs on a timer.
 */
async function publishWorkbookName(): Promise<void> {
  const name = await readWorkbookName();
  if (name && client?.connected) client.sendWorkbookInfo(name);
}

/* -------------------------------------------------------------------------- */
/* Small helpers                                                               */
/* -------------------------------------------------------------------------- */

function renderState(state: string): void {
  connectedFiles.setConnected(state === "open");
  els.statusLabel.textContent = ({ open: `Connected to ${activeHost().app}`, connecting: "Connecting…", closed: "Reconnecting…", idle: `Waiting for ${activeHost().app}` } as Record<string, string>)[state] ?? state;
  els.status.dataset.state = state;
  els.send.disabled = state !== "open" || els.send.dataset.stopping === "true";
  if (state === "closed" && currentRequestId) {
    autoQueueAllowed = false;
    holdPendingPrompts();
    // The bridge dropped mid-turn; the running request can no longer finish, so free the UI.
    usagePanel.reset();
    renderUsage();
    currentRequestId = null;
    live = null;
    setBusy(false);
    showToast("Connection to the bridge was lost; the running request was stopped. Reconnecting…");
  }
}

/** The send control becomes a red Stop button while a turn is running. */
function setBusy(busy: boolean): void {
  if (busy) autoQueueAllowed = true;
  if (!busy) clearThinking();
  cancelRequested = false;
  promptDispatched = false;
  els.send.dataset.mode = busy ? "stop" : "send";
  els.send.dataset.stopping = "false";
  els.send.disabled = els.status.dataset.state !== "open";
  els.sendLabel.textContent = busy ? "Stop" : "Send";
  els.send.setAttribute("aria-label", busy ? "Stop response" : "Send message");
  els.send.title = busy ? "Stop response" : "Send message";
  els.status.dataset.busy = String(busy);
  updateSendControl();
}

function abandonRequest(): void {
  autoQueueAllowed = false;
  if (currentRequestId) client.cancel(currentRequestId);
  holdPendingPrompts();
  pendingPrompts.clear();
  currentRequestId = null;
  requestSession = null;
  usagePanel.reset();
  setBusy(false);
}

function updateSendControl(): void {
  const running = currentRequestId !== null;
  const hasPrompt = !!els.input.value.trim() || pendingAttachments.length > 0;
  const stop = running && !hasPrompt;
  els.send.dataset.mode = stop ? "stop" : "send";
  const label = stop ? "Stop response" : running ? "Queue message" : "Send message";
  els.sendLabel.textContent = stop ? "Stop" : running ? "Queue" : "Send";
  els.send.setAttribute("aria-label", label);
  els.send.title = label;
  byId("stop-button").hidden = !running || !hasPrompt;
}

function scrollToEnd(): void {
  els.messages.scrollTop = els.messages.scrollHeight;
}

function summarize(value: unknown): string {
  try {
    const text = JSON.stringify(value);
    return !text || text === "{}" ? "" : text.length > 70 ? `${text.slice(0, 67)}...` : text;
  } catch {
    return "";
  }
}

