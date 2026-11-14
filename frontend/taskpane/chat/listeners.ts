import { byId } from "../ui/dom.js";
import type { ProviderId } from "../../../protocol/src/index.js";
import { confirmAction } from "../ui/confirm-action.js";
import { PROVIDER_NAMES } from "../ui/provider-picker.js";
import { showToast } from "../ui/toast.js";
import { setupUsagePopover } from "../ui/usage-popover.js";
import type { PermissionMode } from "../permissions.js";
import { atomicBadgeDeletion, fileMentionRanges } from "../ui/file-mentions.js";
import { slashBadgeRange } from "../ui/slash.js";
import { filesFromClipboard } from "../ui/transfer-files.js";
import { els } from "../ui/elements.js";
import { addTemplate, backFromPanel, chooseLocalFiles, clearHistory, closeFloating, closeMenu, continueViewedSession, deleteViewedSession, exportUsage, mcpSettings, navigation, openFloating, openPanel, refreshUsage, renderTemplates, saveModelChoice, selectProvider, setupFileDrop, showPanel, slash, store, updateMentionMenu, userSkills } from "./services.js";
import { persistNow, startNewChat } from "./chats.js";
import { renderComposerPreview, selectPermissionMode, syncAttachmentsWithInput, syncSelectedSlashToken } from "./composer.js";
import { app, pendingAttachments } from "./state.js";
import { stopCurrentTurn, submit } from "./turn.js";

/** Hooks every button, menu and the message box of the pane up to the chat. */
export function bindListeners(): void {
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
    void navigation.goToLink({
      ...(link.dataset.file ? { file: link.dataset.file } : {}),
      ...(link.dataset.sheet ? { sheet: link.dataset.sheet } : {}),
      address: link.dataset.address,
    });
  });
  byId("menu-login").addEventListener("click", () => showPanel("login"));
  byId("menu-logout").addEventListener("click", async () => {
    const button = byId("menu-logout");
    if (!await confirmAction(button.parentElement!, `Sign out of ${PROVIDER_NAMES[app.provider]}? You will need to sign in again to use it.`, "Sign out")) return;
    try {
      await app.client.providerLogout(app.provider);
      app.client.setProvider(app.provider);
    } catch (error) {
      showToast(error instanceof Error ? error.message : "Could not sign out.");
    }
  });
  byId("provider-login-button").addEventListener("click", () => showPanel("login"));
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
  els.permissionMode.value = app.permissionMode;
  app.permissionPicker?.refresh();
  els.permissionMode.addEventListener("change", () => selectPermissionMode(els.permissionMode.value as PermissionMode));
  els.modelSelect.addEventListener("change", saveModelChoice);
  els.send.addEventListener("click", () => {
    if (app.currentRequestId && !els.input.value.trim() && !pendingAttachments.length) stopCurrentTurn();
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
    const command = slashBadgeRange(text, app.selectedSlashToken);
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
}
