import type { UsageSummary, TokenUsage, ProviderEvent, ProviderId, ProviderConnectionStatus, ToolResult } from "../../protocol/src/index.js";
import { BridgeClient, type ModelCatalog, type SavedReference, type SkillTemplate, type UploadedAttachment, type WorkspaceEntry, type WorkspaceRoot } from "./bridge-client.js";
import type { PermissionMode } from "./permissions.js";
import { WRITE_TOOLS, executeTool } from "./executor/index.js";
import { formatDuration, usageCsv } from "./ui/usage.js";
import { setupUsagePopover } from "./ui/usage-popover.js";
import { describeChange } from "./ui/changes.js";
import { setupInputResizer } from "./ui/input-resizer.js";
import { renderMarkdown } from "./ui/markdown.js";
import { createSelfUpdater } from "./ui/self-update.js";
import { setupProviderLogin } from "./ui/provider-login.js";
import { setupComposerMenu } from "./ui/composer-menu.js";
import { setupTheme } from "./ui/theme.js";
import { setupSavedReferences } from "./ui/saved-references.js";
import { setupUserSkills } from "./ui/user-skills.js";
import {
  SessionStore,
  buildContextPreface,
  formatWhen,
  newSession,
  type StoredChange,
  type StoredMessage,
  type StoredSession,
  type StoredTool,
} from "./ui/sessions.js";
import { createSlashMenu } from "./ui/slash-menu.js";
import { leadingCommand, slashBadgeRange } from "./ui/slash.js";
import { activeFileMention, atomicBadgeDeletion, fileMentionRanges } from "./ui/file-mentions.js";

const els = {
  usageStats: byId<HTMLElement>("usage-stats"),
  usageBreakdown: byId<HTMLElement>("usage-breakdown"),
  usageCredits: byId<HTMLElement>("usage-credits"),
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
    skills: byId<HTMLElement>("panel-skills"),
    login: byId<HTMLElement>("panel-login"),
    changes: byId<HTMLElement>("panel-changes"),
    history: byId<HTMLElement>("panel-history"),
    references: byId<HTMLElement>("panel-references"),
    templates: byId<HTMLElement>("panel-templates"),
    session: byId<HTMLElement>("panel-session"),
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
/** The provider process starts empty after a reload or when a past chat is continued; replay the transcript once. */
let needsContext = false;
let currentRequestId: string | null = null;
let cancelRequested = false;
let promptDispatched = false;
let live: { element: HTMLElement; message: StoredMessage } | null = null;
const liveTools = new Map<string, { card: HTMLElement; tool: StoredTool }>();
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
});
const savedReferences = setupSavedReferences(() => client);
const userSkills = setupUserSkills(() => client, () => slash.reload());

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  client = new BridgeClient({
    onState: renderState,
    onHello: ({ addinBuildId }) => {
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
    onError: (message) => append({ kind: "error", text: message }),
    onToolExecuted: recordChange,
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
  els.menu.querySelectorAll<HTMLButtonElement>("[data-provider]").forEach((button) =>
    button.addEventListener("click", () => selectProvider(button.dataset.provider as ProviderId)),
  );
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
  els.menuReferences.addEventListener("click", () => showPanel("references"));
  els.menuTemplates.addEventListener("click", () => showPanel("templates"));
  byId("menu-skills").addEventListener("click", () => showPanel("skills"));
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
    if (currentRequestId) stopCurrentTurn();
    else void submit();
  });
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
    case "tool_group": {
      const group = document.createElement("details");
      group.className = "tool-group";
      const summary = document.createElement("summary");
      summary.textContent = toolGroupLabel(message.tools?.length ?? 0);
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

function renderToolCard(tool: StoredTool): HTMLElement {
  const card = document.createElement("div");
  card.className = "tool";
  card.innerHTML = '<span class="mark"></span><span class="name"></span><span class="args"></span>';
  card.querySelector(".name")!.textContent = tool.name;
  paintTool(card, tool);
  return card;
}

function toolGroupLabel(count: number): string {
  return `${count} tool${count === 1 ? "" : "s"} called`;
}

function bubble(kind: "user" | "assistant" | "error"): HTMLDivElement {
  const element = document.createElement("div");
  element.className = `bubble ${kind}`;
  return element;
}

function paintTool(card: HTMLElement, tool: StoredTool): void {
  card.dataset.state = tool.state;
  card.querySelector(".args")!.textContent = tool.error ?? tool.args;
  if (tool.error) card.title = tool.error;
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

async function submit(): Promise<void> {
  const prompt = els.input.value.trim() || (pendingAttachments.length ? "Please inspect the attached files." : "");
  if (!prompt || currentRequestId || !client.connected) return;
  const model = els.modelSelect.value || null;
  const effort = els.effortSelect.value as "low" | "medium" | "high";
  saveModelChoice();
  if (!pendingAttachments.length && /(?:^|\s)@[^\n]*\.(?:xlsx|xlsm|xltx|xltm|png|jpe?g|gif|webp|md|txt|json|csv|ya?ml|zip)(?=\s|$)/i.test(prompt)) {
    append({ kind: "error", text: "The @ file is only text. Choose it from a workspace or the local file picker so I can read it." });
    updateMentionMenu();
    return;
  }
  slash.close();
  if (prompt === "/clear") {
    els.input.value = "";
    startNewChat();
    return;
  }

  const requestId = randomId();
  currentRequestId = requestId;
  setBusy(true);
  els.input.value = "";
  renderComposerPreview();
  showPanel(null);
  const context = needsContext || (activeModel !== undefined && activeModel !== model) || (activeEffort !== undefined && activeEffort !== effort)
    ? buildContextPreface(session) : undefined;
  needsContext = false;
  append({ kind: "user", text: prompt, attachments: pendingAttachments.map((item) => ({ name: item.name })) });
  showThinking();
  persistNow();

  const result = await executeTool("get_selection", {});
  const selection = result.ok ? (result.data as never) : null;
  if (cancelRequested) return finishLocalCancel(requestId);
  try {
    const attachmentIds = pendingAttachments.flatMap((item) => item.reference.id ? [item.reference.id] : []);
    const referencePaths = pendingAttachments.flatMap((item) => item.reference.id ? [] : [item.reference.path]);
    if (cancelRequested) return finishLocalCancel(requestId);
    client.sendPrompt(requestId, prompt, selection, context, attachmentIds, model, referencePaths, effort, session.id);
    clearAttachment();
    promptDispatched = true;
    activeModel = model;
    activeEffort = effort;
  } catch (error) {
    if (cancelRequested) return finishLocalCancel(requestId);
    clearThinking();
    currentRequestId = null;
    setBusy(false);
    els.input.value = prompt;
    renderComposerPreview();
    append({ kind: "error", text: error instanceof Error ? error.message : String(error) });
  }
}

function stopCurrentTurn(): void {
  if (!currentRequestId || cancelRequested) return;
  cancelRequested = true;
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
  const last = session.messages[session.messages.length - 1];
  if (last?.kind === "user") {
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
      button.textContent = `${isFolder ? "📁" : "📄"}  ${entry.name}${!workspaceStack.length && "isDefault" in entry && entry.isDefault ? "  · current workbook" : ""}${!supported ? "  · unsupported" : ""}`;
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
    if (!workspaceStack.length && workspaceDefaultError) menu.append(workspaceNote("Workbook folder unavailable. Choose local files below."));
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
        button.textContent = `📄  ${ref.name}${ref.available ? "" : "  · file not found"}`;
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
  pick.textContent = "Choose local files…";
  pick.addEventListener("click", openLocalFilePicker);
  actions.append(pick);
  menu.append(actions);
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
  switch (event.type) {
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
      if (!live) live = startLive("assistant");
      live.message.text += event.text;
      live.element.innerHTML = renderMarkdown(live.message.text);
      scrollToEnd();
      schedulePersist();
      return;
    case "message_completed":
      if (!live) live = startLive("assistant");
      live.message.text = event.text;
      live.element.innerHTML = renderMarkdown(event.text);
      live = null; // the next message (after a tool call) gets its own bubble
      scrollToEnd();
      schedulePersist();
      return;
    case "tool_started": {
      live = null;
      const tool: StoredTool = { name: event.tool, args: summarize(event.args), state: "running", error: null };
      if (!liveToolGroup) {
        const message: StoredMessage = { kind: "tool_group", text: "", tools: [] };
        const element = append(message) as HTMLDetailsElement;
        element.open = true;
        liveToolGroup = { element, message };
      }
      liveToolGroup.message.tools!.push(tool);
      liveToolGroup.element.querySelector("summary")!.textContent = toolGroupLabel(liveToolGroup.message.tools!.length);
      const card = renderToolCard(tool);
      liveToolGroup.element.querySelector(".tool-list")!.append(card);
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
      persistNow();
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
    ? `Input: ${input.toLocaleString()} · Cached: ${cached.toLocaleString()} · Output: ${output.toLocaleString()}`
    : "Provider has not reported token usage";
  els.usageBreakdown.title = "Cached tokens are included in input, not added again. Input includes instructions, tools, workbook context and conversation history.";
  const creditTurns = (summary?.creditReportedTurns ?? 0) + (running && turnUsage?.copilotCredits !== undefined ? 1 : 0);
  const credits = (summary?.copilotCredits ?? 0) + (running ? turnUsage?.copilotCredits ?? 0 : 0);
  els.usageCredits.hidden = provider !== "copilot" && creditTurns === 0;
  els.usageCredits.textContent = creditTurns
    ? `Copilot credits: ≈ ${credits.toFixed(2)}${creditTurns < turns ? " (partial)" : ""}`
    : "Copilot credits: not reported";
  els.usageCredits.title = "Provider-reported AI Credits from Copilot /usage, rounded by the CLI. Separate from token counts; never estimated from tokens.";
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

function startLive(kind: "assistant"): { element: HTMLElement; message: StoredMessage } {
  const message: StoredMessage = { kind, text: "" };
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
  if (currentRequestId) client.cancel(currentRequestId);
  clearAttachment();
  persistNow();
  session = newSession(provider);
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
  clearAttachment();
  session = stored;
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
  els.menu.querySelectorAll<HTMLButtonElement>("[data-provider]").forEach((button) => {
    button.setAttribute("aria-checked", String(button.dataset.provider === provider));
  });
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

function openPanel(): "changes" | "history" | "session" | "templates" | "login" | "skills" | "references" | null {
  for (const name of ["changes", "history", "session", "templates", "login", "skills", "references"] as const) if (!els.panels[name].hidden) return name;
  return null;
}

function showPanel(name: "changes" | "history" | "session" | "templates" | "login" | "skills" | "references" | null): void {
  closeMenu();
  for (const key of ["changes", "history", "session", "templates", "login", "skills", "references"] as const) els.panels[key].hidden = key !== name;
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

function renderHistory(): void {
  persistNow(); // so the current chat is listed too
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

function watchSelection(): void {
  let timer: number | undefined;
  const refresh = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address,rowCount,columnCount");
        await context.sync();
        const cells = range.rowCount * range.columnCount;
        els.chip.textContent = cells > 1 ? `${range.address} · ${cells} cells` : range.address;
      });
    } catch (error) {
      // A selected shape/chart cannot be read as a range: say so instead of "No selection".
      els.chip.textContent = (error as { code?: string })?.code === "InvalidSelection" ? "Object selected (shape / chart)" : "No selection";
    }
  };
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => void refresh(), 250);
  });
  void refresh();
}

/* -------------------------------------------------------------------------- */
/* Small helpers                                                               */
/* -------------------------------------------------------------------------- */

function renderState(state: string): void {
  els.statusLabel.textContent = ({ open: "Connected to Excel", connecting: "Connecting…", closed: "Reconnecting…", idle: "Waiting for Excel" } as Record<string, string>)[state] ?? state;
  els.status.dataset.state = state;
  els.send.disabled = state !== "open" || els.send.dataset.stopping === "true";
  if (state === "closed" && currentRequestId) {
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
