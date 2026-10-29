import { watchShapeSelectionEvents } from "./executor/selection-events.js";
import { getSelectionLabel } from "./executor/read.js";
import { monitorSelection } from "./ui/selection.js";
import type { UsageSummary, TokenUsage, ProviderEvent, ProviderId, ProviderConnectionStatus, ToolResult } from "../../protocol/src/index.js";
import { BridgeClient, type ModelCatalog, type SavedReference, type SkillTemplate, type UploadedAttachment, type WorkspaceEntry, type WorkspaceRoot } from "./bridge-client.js";
import type { PermissionMode } from "./permissions.js";
import { WRITE_TOOLS, executeTool } from "./executor/index.js";
import { formatCredits, formatDuration, usageCsv } from "./ui/usage.js";
import { setupUsagePopover } from "./ui/usage-popover.js";
import { describeChange } from "./ui/changes.js";
import { setupInputResizer } from "./ui/input-resizer.js";
import { renderMarkdown } from "./ui/markdown.js";
import { buildPreviewDocument, renderArtifactCard, type StoredArtifact } from "./ui/artifact.js";
import { BrowserView } from "./ui/browser-view.js";
import { createSelfUpdater } from "./ui/self-update.js";
import { setupProviderLogin } from "./ui/provider-login.js";
import { setupToolSettings } from "./ui/tool-settings.js";
import { setupComposerMenu } from "./ui/composer-menu.js";
import { setupTheme } from "./ui/theme.js";
import { setupSavedReferences } from "./ui/saved-references.js";
import { setupUserSkills } from "./ui/user-skills.js";
import { showUserInput } from "./ui/user-input.js";
import { setupSettings } from "./ui/settings.js";
import { renderToolCard, paintTool, formatToolValue, toolChevron } from "./ui/tool-view.js";
import { exposedToolNames } from "../../protocol/src/index.js";
import {
  SessionStore,
  buildContextPreface,
  bindProviderSession,
  formatWhen,
  newSession,
  type StoredChange,
  type StoredMessage,
  type StoredSession,
  type StoredTool,
} from "./ui/sessions.js";
import { createSlashMenu } from "./ui/slash-menu.js";
import { leadingCommand, slashBadgeRange } from "./ui/slash.js";
import { parseNativeCommand, type ProviderCommand } from "../../protocol/src/commands.js";
import { activeFileMention, atomicBadgeDeletion, fileMentionRanges } from "./ui/file-mentions.js";

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

const PROVIDER_KEY = "excelPartner.provider";
const MODEL_KEY = "excelPartner.models.v1";
const PERMISSION_MODE_KEY = "excelPartner.permissionMode";
const NARROW_BELOW = 330;
/** A chat younger than this is reopened automatically after a reload. */
const RESTORE_WITHIN_MS = 12 * 3_600_000;

const PROVIDER_NAMES: Record<ProviderId, string> = { claude: "Claude Code", copilot: "GitHub Copilot", codex: "Codex", opencode: "OpenCode" };
/** Short label used in the model picker. */
const PROVIDER_SHORT_NAMES: Record<ProviderId, string> = { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode" };
/** CLI command that signs the provider in on the bridge machine. */
const PROVIDER_LOGIN_COMMANDS: Record<ProviderId, string> = {
  claude: "claude",
  copilot: "copilot login",
  codex: "codex login",
  opencode: "opencode auth login",
};
/** Hint shown under the manual login command. */
const PROVIDER_LOGIN_HINTS: Record<ProviderId, string> = {
  claude: "Open a terminal, run the command below and use /login. To install the CLI: npm install -g @anthropic-ai/claude-code",
  copilot: "",
  codex: "Open a terminal and run the command below. To install the CLI: npm install -g @openai/codex",
  opencode: "Open a terminal and run the command below. To install the CLI: npm install -g opencode-ai",
};
const MODEL_ID = /^[a-zA-Z0-9][a-zA-Z0-9._/-]{0,79}$/;

function loadModels(): Partial<Record<ProviderId, string>> {
  try {
    const parsed = JSON.parse(localStorage.getItem(MODEL_KEY) ?? "{}") as Record<string, unknown>;
    return Object.fromEntries(Object.entries(parsed).filter(([key, value]) =>
      (key === "claude" || key === "copilot" || key === "codex" || key === "opencode") && typeof value === "string" && MODEL_ID.test(value),
    )) as Partial<Record<ProviderId, string>>;
  } catch { return {}; }
}

const store = new SessionStore(localStorage);
let client: BridgeClient;
let provider: ProviderId = loadProvider();
const models = loadModels();
const modelCatalogs = new Map<ProviderId, ModelCatalog>();
const modelErrors = new Map<ProviderId, boolean>();
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
const workbookToolNames = new Set<string>(exposedToolNames(true, { playwright: true, webSearch: true }));
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
byId("menu-browser").addEventListener("click", () => { showPanel("browser"); });
const previewFrame = byId<HTMLIFrameElement>("preview-frame");
const previewSource = byId<HTMLElement>("preview-source");
const previewToggle = byId<HTMLButtonElement>("preview-toggle");
const previewCopy = byId<HTMLButtonElement>("preview-copy");
let previewing: StoredArtifact | null = null;
let liveToolGroup: { element: HTMLDetailsElement; message: StoredMessage } | null = null;
let thinking: HTMLElement | null = null;
const pendingAttachments: Array<{ reference: UploadedAttachment; name: string }> = [];
let workspaceRoots: WorkspaceRoot[] | null = null;
let mentionReferences: SavedReference[] = [];
let workspaceStack: { path: string; name: string }[] = [];
let workspaceEntries: WorkspaceEntry[] = [];
let workspaceError: string | null = null;
let workspaceDefaultError: string | null = null;
let workspaceLoading = false;
let workspaceRequest = 0;
let localPickerLoading = false;
let viewing: StoredSession | null = null;
let persistTimer: number | undefined;
let usageSummary: UsageSummary | undefined;
let usageSessionId: string | null = null;
let usageRequestId: string | null = null;
let usageStartedAt: number | null = null;
let turnUsage: TokenUsage | null = null;
let usageTimer: number | undefined;
const providerHints = new Map<ProviderId, string>();
const providerStatuses = new Map<ProviderId, ProviderConnectionStatus>();
const modelLoginRequired = new Set<ProviderId>();
let modelPicker: ReturnType<typeof setupComposerMenu> | undefined;
let permissionPicker: ReturnType<typeof setupComposerMenu> | undefined;
let selectedSlashToken: string | null = null;

const updater = createSelfUpdater({ isBusy: () => currentRequestId !== null });
const loginPanel = setupProviderLogin({
  client: () => client,
  isBusy: () => currentRequestId !== null,
  onSuccess: () => {
    modelLoginRequired.delete("copilot");
    providerStatuses.set("copilot", "ready");
    renderAuthNotice();
    if (provider !== "copilot") selectProvider("copilot");
    else {
      client.setProvider("copilot");
      modelCatalogs.delete("copilot");
      void refreshModelCatalog("copilot");
    }
    showPanel(null);
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
const savedReferences = setupSavedReferences(() => client);
const userSkills = setupUserSkills(() => client, () => slash.reload());
const toolSettings = setupToolSettings(() => provider);
const settings = setupSettings({
  loadInstructions: () => client.systemPrompt(),
  saveInstructions: (content, revision) => client.saveSystemPrompt(content, revision),
  historyCount: () => store.list().length,
  clearHistory: () => els.historyClear.click(), historyClearButton: els.historyClear,
});

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

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
      if (session === owner) questionBubble.innerHTML = renderMarkdown(questionText);
      if (session === owner && requestId === currentRequestId) {
        append({ kind: response.cancelled ? "note" : "user", text: response.cancelled ? "Question skipped or cancelled." : request.questions.map(q =>
          `${q.header || q.question}: ${q.isSecret ? "[Hidden answer]" : response.answers[q.id]!.join(", ")}`).join("\n") });
        if (!signal.aborted) showThinking();
      }
      return response;
    },
    onState: renderState,
    onHello: ({ addinBuildId }) => {
      slash.reload();
      client.setProvider(provider);
      void refreshUsage();
      void refreshModelCatalog(provider);
      if (session.messages.length) needsContext = true;
      activeModel = undefined;
      workspaceRoots = null;
      workspaceStack = [];
      void ensureWorkspaces();
      updater.check(addinBuildId);
    },
    onProviderStatus: ({ provider: id, detection }) => {
      providerStatuses.set(id, detection.status === "ready" && modelLoginRequired.has(id) ? "login_required" : detection.status);
      providerHints.set(id, `${detection.version ?? "not found"} · ${detection.status.replace("_", " ")}`);
      renderProviderMenu();
      renderAuthNotice();
      if (id === provider && detection.status === "error") {
        append({ kind: "error", text: detection.hint ?? `${PROVIDER_NAMES[id]} is ${detection.status}.` });
      }
    },
    onProviderEvent: handleEvent,
    onBrowserFrame: (frame) => { if (frame.requestId === currentRequestId) browserView.push(frame); },
    onError: (message) => append({ kind: "error", text: message }),
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
      recordChange(tool, args, result);
      if (result.ok && ["add_shape", "add_connector", "add_line", "update_shape", "delete_shapes", "execute_office_js", "manage_sheet", "insert_sheets_from_file"].includes(tool)) {
        refreshSelectionSubscriptions?.();
      }
    },
  });
  client.setPermissionMode(permissionMode);
  client.connect();

  setupTheme(byId<HTMLSelectElement>("theme-select"));
  modelPicker = setupComposerMenu(els.modelSelect, byId<HTMLButtonElement>("model-button"), byId("model-menu"), (label) => `${PROVIDER_SHORT_NAMES[provider]} · ${label}`);
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
    els.menu.hidden ? openMenu() : closeMenu();
  });
  document.addEventListener("click", (event) => {
    if (!els.menu.hidden && !els.menu.contains(event.target as Node)) closeMenu();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (!els.menu.hidden) closeMenu();
    else if (!slash.isOpen && openPanel() !== null) showPanel(null);
  });
  byId<HTMLSelectElement>("provider-select").addEventListener("change", event => selectProvider((event.target as HTMLSelectElement).value as ProviderId));
  els.menuChanges.addEventListener("click", () => showPanel("changes"));
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
  byId("tools-button").addEventListener("click", () => {
    showPanel("settings");
    byId("settings-tools-section").scrollIntoView({ block: "nearest" });
  });
  els.menuNew.addEventListener("click", () => {
    closeMenu();
    startNewChat();
  });

  // Panels
  document.querySelectorAll<HTMLButtonElement>("[data-close]").forEach((button) => button.addEventListener("click", () => showPanel(null)));
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
  els.input.addEventListener("blur", () => window.setTimeout(() => {
    slash.close();
    if (!els.mentionMenu.contains(document.activeElement)) els.mentionMenu.hidden = true;
  }, 120));
  window.addEventListener("pagehide", persistNow);

  // Reopen the last chat after a reload (e.g. the pane updated itself).
  const previous = store.latest(RESTORE_WITHIN_MS);
  if (previous) loadSession(previous);
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
      element.innerHTML = renderMarkdown(message.text);
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

async function submit(queued?: PendingPrompt): Promise<void> {
  const prompt = queued?.prompt ?? (els.input.value.trim() || (pendingAttachments.length ? "Please inspect the attached files." : ""));
  if (!prompt || !client.connected) return;
  if (prompt.length > 32_000) { append({ kind: "error", text: "Messages must be 32,000 characters or fewer." }); return; }
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
    } catch (error) { append({ kind: "error", text: error instanceof Error ? error.message : String(error) }); return; }
    finally { resolvingCommand = false; }
  }
  const turnAttachments = queued?.attachments ?? pendingAttachments.slice();
  if (!turnAttachments.length && /(?:^|\s)@[^\n]*\.(?:xlsx|xlsm|xltx|xltm|png|jpe?g|gif|webp|md|txt|json|csv|ya?ml|zip)(?=\s|$)/i.test(prompt)) {
    append({ kind: "error", text: "The @ file is only text. Choose it from a workspace or the local file picker so I can read it." });
    updateMentionMenu();
    return;
  }
  if (currentRequestId) {
    if (queued || cancelRequested) return;
    if (pendingPrompts.size >= 20) { append({ kind: "error", text: "Wait for pending messages to be sent before adding more." }); return; }
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
    } catch (error) { pendingPrompts.delete(id); append({ kind: "error", text: error instanceof Error ? error.message : String(error) }); }
    return;
  }
  if (!session.identityVersion && !session.providerSessionId && session.messages.some(message => message.kind === "user")) {
    append({ kind: "error", text: "This legacy chat has no saved CLI session ID, so it cannot be resumed reliably. Its history is preserved. Start a new chat to create a native session." });
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
  browserView.reset();
  setBusy(true);
  if (!queued) els.input.value = "";
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
  } else append({ kind: "user", text: prompt, attachments: turnAttachments.map((item) => ({ name: item.name })) });
  if (!queued) clearAttachment();
  showThinking();
  persistNow();

  const result = nativeCommand ? null : await executeTool("get_selection", {});
  if (currentRequestId !== requestId || requestSession !== session) return;
  const selection = result?.ok ? (result.data as never) : null;
  if (cancelRequested) return finishLocalCancel(requestId);
  try {
    const attachmentIds = turnAttachments.flatMap((item) => item.reference.id ? [item.reference.id] : []);
    const referencePaths = turnAttachments.flatMap((item) => item.reference.id ? [] : [item.reference.path]);
    if (cancelRequested) return finishLocalCancel(requestId);
    client.sendPrompt(requestId, prompt, selection, context, attachmentIds, model, referencePaths, effort, session.id,
      session.provider === provider ? session.providerSessionId : undefined, toolSettings.value(), nativeCommand || undefined);
    promptDispatched = true;
    activeModel = model;
    activeEffort = effort;
  } catch (error) {
    if (cancelRequested) return finishLocalCancel(requestId);
    clearThinking();
    currentRequestId = null;
    setBusy(false);
    if (queued) { queued.state = "held"; pendingPrompts.set(queued.id, queued); updateDelivery(queued); }
    else els.input.value = prompt;
    renderComposerPreview();
    append({ kind: "error", text: error instanceof Error ? error.message : String(error) });
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

function finishLocalCancel(requestId: string): void {
  if (currentRequestId !== requestId) return;
  clearThinking();
  currentRequestId = null;
  holdPendingPrompts();
  const last = session.messages[session.messages.length - 1];
  if (last?.kind === "user" && pendingPrompts.size === 0) {
    session.messages.pop();
    const bubbles = els.messages.querySelectorAll(".bubble.user");
    bubbles[bubbles.length - 1]?.remove();
    els.input.value = last.text;
    renderComposerPreview();
  }
  needsContext = session.messages.length > 0;
  setBusy(false);
  append({ kind: "note", text: "Stopped." });
  persistNow();
  updater.reloadIfIdle();
}

function updateMentionMenu(): void {
  const before = els.input.value.slice(0, els.input.selectionStart);
  const mention = activeFileMention(before, pendingAttachments.map((item) => item.name));
  const show = !!mention;
  const opening = show && els.mentionMenu.hidden;
  els.mentionMenu.hidden = !show;
  if (show) {
    slash.close();
    void ensureWorkspaces();
    if (opening) void refreshMentionReferences();
    renderWorkspaceMenu();
  }
}

/** Saved reference files are listed in the @ menu; reload them each time it opens so edits show up. */
async function refreshMentionReferences(): Promise<void> {
  try { mentionReferences = await client.listSavedReferences(); }
  catch { mentionReferences = []; }
  renderWorkspaceMenu();
}

async function ensureWorkspaces(): Promise<void> {
  if (workspaceRoots || workspaceLoading) return;
  workspaceLoading = true;
  workspaceError = null;
  renderWorkspaceMenu();
  try {
    const [documentUrl, workbookName] = await Promise.all([currentWorkbookUrl(), currentWorkbookName()]);
    const result = await client.listWorkspaces(documentUrl, workbookName);
    workspaceRoots = result.roots;
    workspaceDefaultError = result.defaultError;
  } catch (error) {
    workspaceError = error instanceof Error ? error.message : String(error);
  } finally {
    workspaceLoading = false;
    renderWorkspaceMenu();
  }
}

async function currentWorkbookUrl(): Promise<string> {
  const direct = Office.context.document.url;
  if (direct) return direct;
  return new Promise((resolve) => {
    const timer = window.setTimeout(() => resolve(""), 2500);
    try {
      Office.context.document.getFilePropertiesAsync((result) => {
        window.clearTimeout(timer);
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value.url ?? "" : "");
      });
    } catch { window.clearTimeout(timer); resolve(""); }
  });
}

async function currentWorkbookName(): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load("name");
      await context.sync();
      return workbook.name;
    });
  } catch { return ""; }
}

async function openWorkspaceFolder(path: string, name: string): Promise<void> {
  const before = els.input.value.slice(0, els.input.selectionStart);
  const mention = activeFileMention(before, pendingAttachments.map((item) => item.name));
  if (mention) {
    const cursor = mention.start + 1;
    els.input.value = `${before.slice(0, cursor)}${els.input.value.slice(els.input.selectionStart)}`;
    els.input.setSelectionRange(cursor, cursor);
    renderComposerPreview();
  }
  workspaceStack.push({ path, name });
  const request = ++workspaceRequest;
  workspaceLoading = true;
  workspaceError = null;
  renderWorkspaceMenu();
  try {
    const listing = await client.listWorkspaceEntries(path);
    if (request === workspaceRequest) workspaceEntries = listing.entries;
  } catch (error) {
    if (request === workspaceRequest) workspaceError = error instanceof Error ? error.message : String(error);
  } finally {
    if (request === workspaceRequest) {
      workspaceLoading = false;
      renderWorkspaceMenu();
    }
  }
}

function renderWorkspaceMenu(): void {
  if (els.mentionMenu.hidden) return;
  const menu = els.mentionMenu;
  menu.replaceChildren();
  const heading = document.createElement("div");
  heading.className = "workspace-heading";
  if (workspaceStack.length) {
    const back = document.createElement("button");
    back.type = "button";
    back.textContent = "←";
    back.title = "Back";
    back.addEventListener("click", () => {
      workspaceStack.pop();
      if (workspaceStack.length) {
        const parent = workspaceStack.pop()!;
        void openWorkspaceFolder(parent.path, parent.name);
      }
      else renderWorkspaceMenu();
    });
    heading.append(back);
  }
  const title = document.createElement("strong");
  title.textContent = workspaceStack.length ? workspaceStack.map((part) => part.name).join(" / ") : "Workspaces";
  title.title = workspaceStack[workspaceStack.length - 1]?.path ?? "Workspaces";
  heading.append(title);
  menu.append(heading);

  if (workspaceLoading) menu.append(workspaceNote("Loading…"));
  else if (workspaceError) menu.append(workspaceNote(workspaceError));
  else {
    const rows = workspaceStack.length ? workspaceEntries : workspaceRoots ?? [];
    const before = els.input.value.slice(0, els.input.selectionStart);
    const query = activeFileMention(before, pendingAttachments.map((item) => item.name))?.query.toLowerCase().trim() ?? "";
    const visible = workspaceStack.length && query ? rows.filter((entry) => entry.name.toLowerCase().includes(query)) : rows;
    if (!visible.length) menu.append(workspaceNote(workspaceStack.length ? "No matching files or folders." : "No workbook folder found. Choose local files below."));
    for (const entry of visible) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "workspace-row";
      button.title = entry.path;
      const isFolder = !workspaceStack.length || ("type" in entry && entry.type === "folder");
      const supported = isFolder || ("supported" in entry && entry.supported);
      setWorkspaceRow(button, `${entry.name}${!workspaceStack.length && "isDefault" in entry && entry.isDefault ? " · current workbook" : ""}${!supported ? " · unsupported" : ""}`, isFolder);
      button.disabled = !supported;
      button.addEventListener("click", () => {
        if (isFolder) void openWorkspaceFolder(entry.path, entry.name);
        else void chooseWorkspaceFile(entry.path);
      });
      if (!workspaceStack.length && "removable" in entry && entry.removable) {
        const line = document.createElement("div");
        line.className = "workspace-root";
        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "workspace-remove";
        remove.textContent = "×";
        remove.title = `Remove workspace ${entry.name}`;
        remove.setAttribute("aria-label", `Remove workspace ${entry.name}`);
        remove.addEventListener("click", () => void removeWorkspace(entry.path));
        line.append(button, remove);
        menu.append(line);
      } else menu.append(button);
    }
    if (!workspaceStack.length && workspaceDefaultError && visible.length) menu.append(workspaceNote("Workbook folder unavailable. Choose local files below."));
  }

  if (!workspaceStack.length && !workspaceLoading && !workspaceError) {
    const before = els.input.value.slice(0, els.input.selectionStart);
    const query = activeFileMention(before, pendingAttachments.map((item) => item.name))?.query.toLowerCase().trim() ?? "";
    const refs = query ? mentionReferences.filter((ref) => `${ref.name} ${ref.path}`.toLowerCase().includes(query)) : mentionReferences;
    if (refs.length) {
      const label = document.createElement("div");
      label.className = "workspace-heading";
      label.textContent = "Reference files";
      menu.append(label);
      for (const ref of refs) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "workspace-row";
        button.title = ref.available ? `${ref.path}${ref.description ? `\n${ref.description}` : ""}` : `File not found: ${ref.path}`;
        setWorkspaceRow(button, `${ref.name}${ref.available ? "" : " · file not found"}`, false);
        button.disabled = !ref.available;
        button.addEventListener("click", () => void chooseWorkspaceFile(ref.path));
        menu.append(button);
      }
    }
  }

  const actions = document.createElement("div");
  actions.className = "workspace-actions";
  const pick = document.createElement("button");
  pick.type = "button";
  setWorkspaceRow(pick, "Choose local files…", false);
  pick.addEventListener("click", openLocalFilePicker);
  actions.append(pick);
  menu.append(actions);
}

function setWorkspaceRow(button: HTMLButtonElement, text: string, folder: boolean): void {
  const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  icon.setAttribute("viewBox", "0 0 20 20");
  icon.setAttribute("aria-hidden", "true");
  icon.classList.add("workspace-icon");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", folder ? "M2.5 6V4.5h6l2 2h7v10h-15z" : "M5 2.5h6l4 4v11H5z M11 2.5v4h4 M7.5 10h5 M7.5 13h5");
  icon.append(path);
  const label = document.createElement("span");
  label.className = "workspace-row-label";
  label.textContent = text;
  button.replaceChildren(icon, label);
}

async function removeWorkspace(path: string): Promise<void> {
  workspaceLoading = true;
  workspaceError = null;
  renderWorkspaceMenu();
  try {
    await client.removeWorkspace(path);
    workspaceRoots = workspaceRoots?.filter((root) => root.path !== path) ?? null;
    workspaceStack = [];
  } catch (error) {
    workspaceError = error instanceof Error ? error.message : String(error);
  } finally {
    workspaceLoading = false;
    renderWorkspaceMenu();
  }
}

function workspaceNote(message: string): HTMLElement {
  const note = document.createElement("div");
  note.className = "slash-empty";
  note.textContent = message;
  return note;
}

async function chooseWorkspaceFile(path: string): Promise<void> {
  workspaceLoading = true;
  renderWorkspaceMenu();
  try {
    const attachment = await client.attachWorkspaceFile(path);
    selectWorkspaceAttachment(attachment);
  } catch (error) {
    workspaceError = error instanceof Error ? error.message : String(error);
  } finally {
    workspaceLoading = false;
    renderWorkspaceMenu();
  }
}

function selectWorkspaceAttachment(attachment: UploadedAttachment): void {
  addAttachment({ reference: attachment, name: attachment.name });
}

function openLocalFilePicker(): void {
  if (localPickerLoading) return;
  const input = document.createElement("input");
  input.type = "file";
  input.multiple = true;
  input.onchange = () => {
    const files = Array.from(input.files ?? []);
    if (files.length) void chooseLocalFiles(files);
  };
  input.click();
}

async function chooseLocalFiles(files: File[]): Promise<void> {
  if (localPickerLoading) return;
  if (pendingAttachments.length + files.length > 10) {
    append({ kind: "error", text: "Choose at most 10 files per message." });
    return;
  }
  localPickerLoading = true;
  els.mentionMenu.hidden = true;
  try {
    for (const file of files) selectWorkspaceAttachment(await client.uploadAttachment(file));
  } catch (error) {
    append({ kind: "error", text: error instanceof Error ? error.message : String(error) });
  } finally {
    localPickerLoading = false;
  }
}

function addAttachment(item: { reference: UploadedAttachment; name: string }): void {
  if (pendingAttachments.length >= 10) {
    append({ kind: "error", text: "Choose at most 10 files per message." });
    return;
  }
  els.mentionMenu.hidden = true;
  const cursor = els.input.selectionStart;
  const before = els.input.value.slice(0, cursor);
  const after = els.input.value.slice(cursor);
  const token = `@${item.name}`;
  const mention = activeFileMention(before, pendingAttachments.map((selected) => selected.name));
  const prefix = mention ? before.slice(0, mention.start) : `${before}${before && !/\s$/.test(before) ? " " : ""}`;
  const separator = !after || !/^\s/.test(after) ? " " : "";
  els.input.value = `${prefix}${token}${separator}${after}`;
  const nextCursor = prefix.length + token.length + 1;
  els.input.setSelectionRange(nextCursor, nextCursor);
  pendingAttachments.push(item);
  renderComposerPreview();
  els.input.focus();
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
        if (event.message) append({ kind: "error", text: event.message });
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
        usageSummary = undefined;
        if (usageRequestId === event.requestId) usageSessionId = session.id;
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
    case "turn_started":
      if (event.requestId !== currentRequestId) return;
      usageSessionId = session.id;
      usageRequestId = event.requestId;
      usageStartedAt = Date.now();
      turnUsage = null;
      window.clearInterval(usageTimer);
      usageTimer = window.setInterval(renderUsage, 1000);
      renderUsage();
      return;
    case "usage_update":
      if (event.requestId === usageRequestId) {
        turnUsage = event.usage;
        renderUsage();
      }
      return;
    case "message_delta":
      if (!event.text || (!live && !event.text.trim())) return;
      if (!live) live = startLive("assistant", event.text);
      else live.message.text += event.text;
      live.element.innerHTML = renderMarkdown(live.message.text);
      scrollToEnd();
      schedulePersist();
      return;
    case "message_completed":
      // Providers can finish an assistant message that only calls a tool.
      // It has no chat text and must not create (or erase) a visible bubble.
      if (!event.text.trim()) { live = null; return; }
      if (!live) live = startLive("assistant", event.text);
      live.message.text = event.text;
      live.element.innerHTML = renderMarkdown(event.text);
      live = null; // the next message (after a tool call) gets its own bubble
      scrollToEnd();
      schedulePersist();
      return;
    case "tool_started": {
      live = null;
      // Workbook calls get their exact result from onToolExecuted. Provider-local
      // tools have a separate native call ID and must not be paired by name.
      if (workbookToolNames.has(event.tool)) return;
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
      if (event.requestId === usageRequestId) {
        window.clearInterval(usageTimer);
        usageStartedAt = null;
        turnUsage = null;
        usageRequestId = null;
      }
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
      sendNextPending();
      updater.reloadIfIdle();
      return;
    case "error":
      clearThinking();
      if (event.code === "login_required") {
        modelLoginRequired.add(provider);
        providerStatuses.set(provider, "login_required");
        renderAuthNotice();
      }
      append({ kind: "error", text: event.message });
      return;
    default:
      return;
  }
}

async function refreshUsage(): Promise<void> {
  const id = session.id;
  if (usageSummary?.sessionId !== id) usageSummary = undefined;
  renderUsage();
  if (!client?.connected) return;
  try {
    const summaries = await client.usageSummaries(id);
    if (session.id !== id) return;
    usageSummary = summaries[0];
    if (usageSummary?.title) {
      session.providerTitle = usageSummary.title;
      session.title = usageSummary.title;
      store.save(session);
    }
    renderUsage();
  } catch {
    if (session.id === id) els.usageStats.textContent = "Usage statistics unavailable";
  }
}

function renderUsage(): void {
  const running = usageSessionId === session.id && usageStartedAt !== null;
  const summary = usageSummary?.sessionId === session.id ? usageSummary : undefined;
  const reported = (summary?.reportedTurns ?? 0) + (running && turnUsage ? 1 : 0);
  const turns = (summary?.turns ?? 0) + (running ? 1 : 0);
  const tokens = (summary?.totalTokens ?? 0) + (running ? turnUsage?.totalTokens ?? 0 : 0);
  const partial = reported > 0 && reported < turns;
  const duration = (summary?.durationMs ?? 0) + (running ? Date.now() - usageStartedAt! : 0);
  els.usageStats.textContent = `Session tokens: ${reported || !turns ? tokens.toLocaleString() : "N/A"}${partial ? " (partial)" : ""} · AI time: ${formatDuration(duration)}`;
  const input = (summary?.inputTokens ?? 0) + (running ? turnUsage?.inputTokens ?? 0 : 0);
  const cached = (summary?.cachedInputTokens ?? 0) + (running ? turnUsage?.cachedInputTokens ?? 0 : 0);
  const output = (summary?.outputTokens ?? 0) + (running ? turnUsage?.outputTokens ?? 0 : 0);
  els.usageBreakdown.textContent = reported || !turns
    ? `Input: ${input.toLocaleString()} · Cache read: ${cached.toLocaleString()} · Output: ${output.toLocaleString()}`
    : "Provider has not reported token usage";
  els.usageBreakdown.title = "Cached tokens are included in input, not added again. Input includes instructions, tools, workbook context and conversation history.";
  const splitTurns = (summary?.cacheSplitReportedTurns ?? 0) + (running && turnUsage?.cacheWriteTokens !== undefined && turnUsage?.uncachedInputTokens !== undefined ? 1 : 0);
  const uncached = (summary?.uncachedInputTokens ?? 0) + (running ? turnUsage?.uncachedInputTokens ?? 0 : 0);
  const cacheWrite = (summary?.cacheWriteTokens ?? 0) + (running ? turnUsage?.cacheWriteTokens ?? 0 : 0);
  els.usageCacheSplit.textContent = splitTurns
    ? `Uncached input: ${uncached.toLocaleString()} · Cache write: ${cacheWrite.toLocaleString()}${splitTurns < turns ? " (partial)" : ""}`
    : "Uncached input / cache write: not reported";
  els.usageCacheSplit.title = "Input = uncached input + cache read + cache write when the full split is available. Both cache buckets are already included in input.";
  const creditTurns = (summary?.creditReportedTurns ?? 0) + (running && turnUsage?.copilotCredits !== undefined ? 1 : 0);
  const credits = (summary?.copilotCredits ?? 0) + (running ? turnUsage?.copilotCredits ?? 0 : 0);
  const exactCreditTurns = (summary?.exactCreditReportedTurns ?? 0) + (running && turnUsage?.copilotCreditNanoAiu !== undefined ? 1 : 0);
  els.usageCredits.hidden = provider !== "copilot" && creditTurns === 0;
  els.usageCredits.textContent = creditTurns
    ? `Copilot credits: ${exactCreditTurns === creditTurns ? "" : "≈ "}${formatCredits(credits)}${creditTurns < turns ? " (partial)" : ""}${exactCreditTurns < creditTurns ? " (includes rounded legacy data)" : ""}`
    : "Copilot credits: not reported";
  els.usageCredits.title = "Original Copilot credits retain integer nano AI units (1 credit = 1,000,000,000 nano units). Old /usage records remain marked rounded. Credits are never estimated from tokens.";
  const originalTurns = (summary?.originalReportedTurns ?? 0) + (running && turnUsage?.source && turnUsage.source !== "copilot_acp" ? 1 : 0);
  const cliSessions = new Set(summary?.providerSessionIds ?? []);
  if (running && turnUsage?.providerSessionId) cliSessions.add(turnUsage.providerSessionId);
  els.usageSource.hidden = cliSessions.size === 0;
  els.usageSource.textContent = `Original provider data: ${originalTurns}/${turns} turns · CLI sessions: ${cliSessions.size}`;
  els.usageSource.title = `Each native CLI session has separate history and accounting. Session: ${[...cliSessions].join(", ") || "not recorded"}. Resuming keeps its ID and previously recorded usage. Background calls after turn completion are outside this report.`;
  els.usageStats.title = `This chat session: ${reported}/${turns} turns with token data. Total includes cached tokens. AI time includes provider/tool/approval waits while processing; idle time is excluded.`;
}

async function exportUsage(): Promise<void> {
  els.usageExport.disabled = true;
  try {
    persistNow();
    const summaries = await client.usageSummaries();
    const titles = new Map(store.list().map((item) => [item.id, item.title]));
    const blob = new Blob([usageCsv(summaries, titles)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `excel-partner-usage-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    append({ kind: "error", text: `Could not export usage: ${error instanceof Error ? error.message : String(error)}` });
  } finally {
    els.usageExport.disabled = false;
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

function clearThinking(): void {
  thinking?.remove();
  thinking = null;
}

function startLive(kind: "assistant", text: string): { element: HTMLElement; message: StoredMessage } {
  const message: StoredMessage = { kind, text };
  session.messages.push(message);
  return { element: renderMessage(els.messages, message), message };
}

/* -------------------------------------------------------------------------- */
/* Chats: new, restore, persist                                                */
/* -------------------------------------------------------------------------- */

function schedulePersist(): void {
  window.clearTimeout(persistTimer);
  persistTimer = window.setTimeout(persistNow, 500);
}

function persistNow(): void {
  window.clearTimeout(persistTimer);
  session.updatedAt = Date.now();
  store.save(session);
}

function startNewChat(): void {
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
  showPanel(null);
}

/** Make a stored chat the current one (its messages are shown in the chat area). */
function loadSession(stored: StoredSession): void {
  slash.reload();
  abandonRequest();
  clearAttachment();
  if (["claude", "copilot", "codex", "opencode"].includes(stored.provider) && stored.provider !== provider) {
    provider = stored.provider as ProviderId;
    localStorage.setItem(PROVIDER_KEY, provider);
    if (client?.connected) client.setProvider(provider);
    renderProviderMenu();
    renderAuthNotice();
    renderModelPicker();
    void refreshModelCatalog(provider);
  }
  session = stored;
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
/* Settings menu + provider                                                    */
/* -------------------------------------------------------------------------- */

function loadProvider(): ProviderId {
  const saved = localStorage.getItem(PROVIDER_KEY);
  return saved === "claude" || saved === "copilot" || saved === "codex" || saved === "opencode" ? saved : "copilot";
}

function openMenu(): void {
  renderProviderMenu();
  els.menu.hidden = false;
  els.menuButton.setAttribute("aria-expanded", "true");
}

function selectPermissionMode(mode: PermissionMode): void {
  if (mode !== "ask" && mode !== "auto") return;
  permissionMode = mode;
  localStorage.setItem(PERMISSION_MODE_KEY, mode);
  client.setPermissionMode(mode);
  els.permissionMode.value = mode;
  permissionPicker?.refresh();
}

function closeMenu(): void {
  els.menu.hidden = true;
  els.menuButton.setAttribute("aria-expanded", "false");
}

function renderProviderMenu(): void {
  toolSettings.refreshProvider();
  byId<HTMLSelectElement>("provider-select").value = provider;
  els.providerHint.textContent = providerHints.get(provider) ?? "";
  byId("menu-login").textContent = `Sign in to ${PROVIDER_NAMES[provider]}`;
}

function renderAuthNotice(): void {
  const status = providerStatuses.get(provider);
  byId("auth-notice").hidden = status !== "login_required" && status !== "not_installed";
  byId("auth-notice-title").textContent = status === "not_installed" ? `Set up ${PROVIDER_NAMES[provider]}` : `Sign in to ${PROVIDER_NAMES[provider]}`;
  byId("auth-notice-detail").textContent = status === "not_installed" ? "Install the CLI on the bridge machine." : "Connect your account to start chatting.";
  byId("provider-login-button").textContent = status === "not_installed" ? "Set up ↗" : "Sign in ↗";
}

function selectProvider(next: ProviderId): void {
  closeMenu();
  if (next === provider) return;
  provider = next;
  activeModel = undefined;
  localStorage.setItem(PROVIDER_KEY, next);
  renderProviderMenu();
  renderAuthNotice();
  renderModelPicker();
  // A provider keeps its own conversation: switching starts a fresh chat (the old one stays in History).
  startNewChat();
  client.setProvider(next);
  void refreshModelCatalog(next);
}

function renderModelPicker(): void {
  const catalog = modelCatalogs.get(provider);
  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = "Default";
  els.modelSelect.replaceChildren(defaultOption, ...(catalog?.models ?? []).map((model) => {
    const option = document.createElement("option");
    option.value = model.id;
    option.textContent = model.name;
    return option;
  }));
  if (!catalog) {
    const info = document.createElement("option");
    info.disabled = true;
    info.textContent = modelErrors.get(provider) ? "Models unavailable" : "Loading models…";
    els.modelSelect.append(info);
  }
  const saved = models[provider];
  els.modelSelect.value = saved && catalog?.models.some((model) => model.id === saved) ? saved : "";
  els.modelSelect.title = `${PROVIDER_NAMES[provider]} model${catalog?.defaultModelId ? `; default: ${catalog.defaultModelId}` : ""}`;
  modelPicker?.refresh();
}

async function refreshModelCatalog(forProvider: ProviderId): Promise<void> {
  try {
    const catalog = await client.listModels(forProvider);
    modelLoginRequired.delete(forProvider);
    if (providerStatuses.get(forProvider) === "login_required") providerStatuses.set(forProvider, "ready");
    renderAuthNotice();
    modelCatalogs.set(forProvider, catalog);
    modelErrors.delete(forProvider);
    if (provider === forProvider) renderModelPicker();
  } catch (error) {
    modelErrors.set(forProvider, true);
    if ((error as { code?: string }).code === "login_required") {
      modelLoginRequired.add(forProvider);
      providerStatuses.set(forProvider, "login_required");
      renderAuthNotice();
    }
    if (provider === forProvider) renderModelPicker();
  }
}

function saveModelChoice(): void {
  const model = els.modelSelect.value;
  if (model) models[provider] = model;
  else delete models[provider];
  localStorage.setItem(MODEL_KEY, JSON.stringify(models));
}

/* -------------------------------------------------------------------------- */
/* Panels: changes, history, past chat                                         */
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

type PanelName = keyof typeof els.panels;
function openPanel(): PanelName | null {
  for (const name of Object.keys(els.panels) as PanelName[]) if (!els.panels[name].hidden) return name;
  return null;
}

function showPanel(name: PanelName | null): void {
  closeMenu();
  for (const key of Object.keys(els.panels) as PanelName[]) els.panels[key].hidden = key !== name;
  document.body.classList.toggle("settings-active", name === "settings");
  // The small view repeats the large one and takes its height, so it steps aside while the large one is open.
  document.body.classList.toggle("browser-active", name === "browser");
  if (name === "settings") { persistNow(); void settings.open(); }
  if (name !== "preview") { previewFrame.srcdoc = ""; previewing = null; }
  if (name === "references") void savedReferences.refresh();
  if (name === "skills") void userSkills.refresh();
  document.body.classList.toggle("login-active", name === "login");
  if (name === "login") {
    byId("login-heading").textContent = `Sign in to ${PROVIDER_NAMES[provider]}`;
    byId("login-description").textContent = provider === "copilot" ? "Finish signing in through your browser, then return to Excel." : "Sign in through the CLI on the bridge machine, then check the connection again.";
    byId("login-device").hidden = provider !== "copilot";
    byId("login-manual").hidden = provider === "copilot";
    byId("login-command").textContent = PROVIDER_LOGIN_COMMANDS[provider];
    byId("login-manual-hint").textContent = PROVIDER_LOGIN_HINTS[provider];
    if (provider === "copilot") loginPanel.open();
    else loginPanel.close();
  } else loginPanel.close();
  byId("tab-chat").setAttribute("aria-pressed", String(name === null));
  byId("tab-changes").setAttribute("aria-pressed", String(name === "changes"));
  byId("tab-history").setAttribute("aria-pressed", String(name === "history" || name === "session"));
  if (name === "changes") renderChanges();
  if (name === "history") renderHistory();
  if (name === "templates") void loadTemplateSkills();
}

async function loadTemplateSkills(): Promise<void> {
  els.templateList.textContent = "Loading skills…";
  try {
    const selected = els.templateSkill.value;
    const skills = await client.listTemplateSkills();
    els.templateSkill.replaceChildren(...skills.map((skill) => {
      const option = document.createElement("option");
      option.value = skill.name;
      option.textContent = skill.name;
      return option;
    }));
    if (skills.some((skill) => skill.name === selected)) els.templateSkill.value = selected;
    els.templateAdd.disabled = skills.length === 0;
    if (skills.length) await renderTemplates();
    else els.templateList.textContent = "No skills found. Add a SKILL.md under the skills folder first.";
  } catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
}

async function renderTemplates(): Promise<void> {
  const skill = els.templateSkill.value;
  if (!skill) return;
  els.templateList.textContent = "Loading templates…";
  try {
    const templates = await client.listSkillTemplates(skill);
    if (skill !== els.templateSkill.value) return;
    els.templateList.replaceChildren(...templates.map((item) => templateRow(skill, item)));
    if (!templates.length) els.templateList.textContent = "No templates for this skill yet.";
  } catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
}

function templateRow(skill: string, item: SkillTemplate): HTMLElement {
  const row = document.createElement("div");
  row.className = "template-row";
  const label = document.createElement("span");
  label.textContent = `${item.name} · ${item.kind} · ${Math.max(1, Math.round(item.size / 1024))} KB`;
  label.title = item.path;
  const remove = document.createElement("button");
  remove.type = "button";
  remove.className = "ghost";
  remove.textContent = "Delete";
  remove.setAttribute("aria-label", `Delete template ${item.name}`);
  remove.addEventListener("click", async () => {
    remove.disabled = true;
    try { await client.deleteSkillTemplate(skill, item.name); await renderTemplates(); }
    catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
  });
  row.append(label, remove);
  return row;
}

async function addTemplate(file: File): Promise<void> {
  const skill = els.templateSkill.value;
  if (!skill) return;
  els.templateAdd.disabled = true;
  try { await client.addSkillTemplate(skill, file); await renderTemplates(); }
  catch (error) { els.templateList.textContent = error instanceof Error ? error.message : String(error); }
  finally { els.templateAdd.disabled = false; }
}

function recordChange(tool: string, args: unknown, result: ToolResult): void {
  const change = describeChange(tool, args, result, WRITE_TOOLS);
  if (!change) return;
  session.changes.push({ at: Date.now(), tool: change.tool, ok: change.ok, summary: change.summary, target: change.target });
  renderChanges();
  schedulePersist();
}

function renderChanges(): void {
  const total = session.changes.length;
  els.changesCount.hidden = total === 0;
  els.changesCount.textContent = String(total);
  const tabCount = byId("tab-changes-count");
  tabCount.hidden = total === 0;
  tabCount.textContent = String(total);
  els.menuDot.hidden = total === 0;
  if (els.panels.changes.hidden) return;

  if (total === 0) {
    els.changesList.innerHTML = '<div class="empty">No changes yet. Edits made to the workbook in this chat will be listed here.</div>';
    return;
  }
  els.changesList.replaceChildren(...[...session.changes].reverse().map((entry) => changeRow(entry)));
}

function changeRow(entry: StoredChange): HTMLElement {
  const row = document.createElement("div");
  row.className = "change";
  row.dataset.ok = String(entry.ok);
  const time = document.createElement("span");
  time.className = "time";
  time.textContent = new Date(entry.at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const summary = document.createElement("span");
  summary.className = "summary";
  summary.textContent = entry.summary;
  row.append(time, summary);
  if (entry.target) {
    const target = entry.target;
    const go = document.createElement("button");
    go.className = "secondary";
    go.textContent = "Go to";
    go.addEventListener("click", () => void goTo(target));
    row.append(go);
  }
  return row;
}

async function goTo(target: { sheet: string; address: string }): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(target.sheet);
      sheet.activate();
      sheet.getRange(target.address).select();
      await context.sync();
    });
  } catch (error) {
    append({ kind: "error", text: `Cannot go to ${target.sheet}!${target.address}: ${error instanceof Error ? error.message : String(error)}` });
  }
}

function renderHistory(refreshTitles = true): void {
  persistNow(); // so the current chat is listed too
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
      open.dataset.current = String(item.id === session.id);
      const title = document.createElement("span");
      title.className = "history-title";
      title.textContent = item.title;
      const meta = document.createElement("span");
      meta.className = "history-meta";
      meta.textContent = [
        item.id === session.id ? "Current" : null,
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
  if (!client?.connected) return;
  try {
    const summaries = await client.usageSummaries();
    let changed = false;
    for (const summary of summaries) {
      if (!summary.title) continue;
      const stored = summary.sessionId === session.id ? session : store.get(summary.sessionId);
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
  persistNow();
  const stored = store.get(id);
  if (!stored) return renderHistory();
  viewing = stored;
  els.sessionTitle.textContent = stored.title;
  els.sessionMeta.textContent = `${PROVIDER_NAMES[stored.provider as ProviderId] ?? stored.provider} · ${new Date(stored.startedAt).toLocaleString()}`;
  els.sessionTranscript.replaceChildren();
  for (const message of stored.messages) renderMessage(els.sessionTranscript, message);
  els.sessionContinue.textContent = stored.id === session.id ? "Back to chat" : "Continue this chat";
  showPanel("session");
  els.sessionTranscript.scrollTop = 0;
}

function continueViewedSession(): void {
  if (!viewing) return;
  if (viewing.id !== session.id) {
    if (currentRequestId) client.cancel(currentRequestId);
    persistNow();
    loadSession(viewing);
  }
  viewing = null;
  showPanel(null);
  els.input.focus();
}

function deleteViewedSession(): void {
  if (viewing) deleteSession(viewing.id);
}

function deleteSession(id: string): void {
  store.delete(id);
  if (id === session.id) {
    abandonRequest();
    // The open chat was removed: carry on in an empty one.
    session = newSession(provider);
    needsContext = false;
    els.messages.replaceChildren();
    renderWelcome();
    renderChanges();
    void refreshUsage();
  }
  viewing = null;
  showPanel("history");
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
  abandonRequest();
  session = newSession(provider);
  needsContext = false;
  els.messages.replaceChildren();
  renderWelcome();
  renderChanges();
  renderHistory();
  void refreshUsage();
}

function resetClearButton(): void {
  els.historyClear.dataset.confirm = "0";
  els.historyClear.textContent = "Delete all";
}

/* -------------------------------------------------------------------------- */
/* Selection chip                                                              */
/* -------------------------------------------------------------------------- */

let refreshSelectionSubscriptions: (() => void) | undefined;

function watchSelection(): void {
  const monitor = monitorSelection({
    read: getSelectionLabel,
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
  const events = watchShapeSelectionEvents({
    onActivated: value => {
      window.clearTimeout(selectionTimer);
      monitor.show(value);
    },
    onSelection: selectionChanged,
    onError: error => { console.warn("Shape selection event registration failed", error); },
  });
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
/* Small helpers                                                               */
/* -------------------------------------------------------------------------- */

function renderState(state: string): void {
  els.statusLabel.textContent = ({ open: "Connected to Excel", connecting: "Connecting…", closed: "Reconnecting…", idle: "Waiting for Excel" } as Record<string, string>)[state] ?? state;
  els.status.dataset.state = state;
  els.send.disabled = state !== "open" || els.send.dataset.stopping === "true";
  if (state === "closed" && currentRequestId) {
    autoQueueAllowed = false;
    holdPendingPrompts();
    // The bridge dropped mid-turn; the running request can no longer finish, so free the UI.
    window.clearInterval(usageTimer);
    usageStartedAt = null;
    turnUsage = null;
    usageRequestId = null;
    renderUsage();
    currentRequestId = null;
    live = null;
    setBusy(false);
    append({ kind: "error", text: "Connection to the bridge was lost; the running request was stopped. Reconnecting…" });
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
  window.clearInterval(usageTimer);
  usageStartedAt = null;
  usageRequestId = null;
  turnUsage = null;
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

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`missing element #${id}`);
  return element as T;
}

function randomId(): string {
  const bytes = new Uint8Array(12);
  crypto.getRandomValues(bytes);
  return [...bytes].map((b) => b.toString(16).padStart(2, "0")).join("");
}
