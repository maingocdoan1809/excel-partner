import type { ProviderId } from "../../../protocol/src/index.js";
import { APP_LOGIN_PROVIDERS } from "../../../protocol/src/provider-login.js";
import { activeHost } from "../host/index.js";
import { byId } from "./dom.js";
import { PROVIDER_NAMES } from "./provider-picker.js";

export type PanelName = "settings" | "skills" | "login" | "changes" | "history" | "references" | "templates" | "session" | "preview" | "browser" | "testing";

/** CLI command that signs the provider in on the bridge machine. */
const PROVIDER_LOGIN_COMMANDS: Record<ProviderId, string> = {
  claude: "office-partner providers login claude",
  copilot: "office-partner providers login copilot",
  codex: "office-partner providers login codex",
  opencode: "office-partner providers login opencode",
};
/** Hint shown under the manual login command. */
const PROVIDER_LOGIN_HINTS: Record<ProviderId, string> = {
  claude: "",
  copilot: "",
  codex: "",
  opencode: "Open a terminal and run the command below. Install a private CLI from System & recovery if needed.",
};

export interface PanelsDeps {
  getProvider(): ProviderId;
  renderHistory(): void;
  renderChanges(): void;
  renderProviderMenu(): void;
  openSettings(returning: boolean): void;
  closePreview(): void;
  refreshReferences(): Promise<void>;
  refreshSkills(): Promise<void>;
  refreshTesting(): Promise<void>;
  loadTemplateSkills(): Promise<void>;
  loginPanel: { open(): void; close(): void };
  elements: {
    panels: Record<PanelName, HTMLElement>;
    menu: HTMLElement;
    menuButton: HTMLButtonElement;
  };
}

/** Full panels, the compact floating menus, and returning from a settings child panel. */
export function setupPanels(deps: PanelsDeps) {
  const els = deps.elements;
  let settingsChild = false;
  let settingsOpener: HTMLElement | null = null;

  function backFromPanel(): void {
    const returning = settingsChild;
    showPanel(returning ? "settings" : null);
    if (returning) settingsOpener?.focus({ preventScroll: true });
    else els.menuButton.focus({ preventScroll: true });
  }

  function openPanel(): PanelName | null {
    for (const name of Object.keys(els.panels) as PanelName[]) if (!els.panels[name].hidden && (!els.panels[name].hasAttribute("popover") || isFloatingOpen(els.panels[name]))) return name;
    return null;
  }

  function isFloatingOpen(element: HTMLElement): boolean {
    return element.matches(":popover-open") || element.dataset.floatingOpen === "true";
  }
  function closeFloating(element: HTMLElement): void {
    if (typeof element.hidePopover === "function") element.hidePopover();
    delete element.dataset.floatingOpen;
    const anchor = document.getElementById(element.dataset.floatingAnchor ?? "");
    anchor?.setAttribute("aria-expanded", "false");
  }

  function openFloating(element: HTMLElement, anchor: HTMLElement): void {
    const wasOpen = isFloatingOpen(element);
    for (const popover of document.querySelectorAll<HTMLElement>(':popover-open, [data-floating-open="true"]')) closeFloating(popover);
    if (wasOpen) return;
    element.classList.add("compact-popover");
    element.setAttribute("popover", "auto");
    element.hidden = false;
    const rect = anchor.getBoundingClientRect();
    const width = Math.min(340, window.innerWidth - 24);
    element.style.width = `${width}px`;
    element.style.left = `${Math.max(12, Math.min(rect.right - width, window.innerWidth - width - 12))}px`;
    element.style.right = "auto";
    const below = rect.bottom < window.innerHeight / 2;
    element.style.top = below ? `${rect.bottom + 8}px` : "auto";
    element.style.bottom = below ? "auto" : `${window.innerHeight - rect.top + 8}px`;
    element.dataset.floatingAnchor = anchor.id;
    if (typeof element.showPopover === "function") element.showPopover();
    else element.dataset.floatingOpen = "true";
    anchor.setAttribute("aria-expanded", "true");
    element.ontoggle = () => anchor.setAttribute("aria-expanded", String(isFloatingOpen(element)));
  }

  function showPanel(name: PanelName | null): void {
    if (name === "history" || name === "changes") {
      if (name === "history") deps.renderHistory(); else deps.renderChanges();
      openFloating(els.panels[name], byId(name === "history" ? "history-button" : "changes-button"));
      return;
    }
    for (const popover of document.querySelectorAll<HTMLElement>(':popover-open, [data-floating-open="true"]')) closeFloating(popover);
    byId("settings-tools-section").append(byId("tools-menu"));
    if (name === "skills") { els.panels.skills.removeAttribute("popover"); els.panels.skills.classList.remove("compact-popover"); els.panels.skills.removeAttribute("style"); }
    const previous = openPanel();
    const returningToSettings = name === "settings" && settingsChild;
    if (previous === "settings" && name !== null && name !== "settings") {
      settingsChild = true;
      settingsOpener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    } else if (name === null || name === "settings") settingsChild = false;
    closeMenu();
    for (const key of Object.keys(els.panels) as PanelName[]) els.panels[key].hidden = key !== name;
    document.body.classList.toggle("settings-active", name === "settings" || settingsChild);
    for (const button of document.querySelectorAll<HTMLButtonElement>("[data-close]")) {
      button.setAttribute("aria-label", settingsChild ? "Back to settings" : "Back to chat");
    }
    // The small view repeats the large one and takes its height, so it steps aside while the large one is open.
    document.body.classList.toggle("browser-active", name === "browser");
    if (name === "settings") deps.openSettings(returningToSettings);
    if (name !== "preview") deps.closePreview();
    if (name === "references") void deps.refreshReferences();
    if (name === "skills") void deps.refreshSkills();
    if (name === "testing") void deps.refreshTesting();
    document.body.classList.toggle("login-active", name === "login");
    if (name === "login") {
      const provider = deps.getProvider();
      byId("login-heading").textContent = `Sign in to ${PROVIDER_NAMES[provider]}`;
      const inApp = APP_LOGIN_PROVIDERS.includes(provider);
      byId("login-description").textContent = inApp ? `Finish signing in through your browser, then return to ${activeHost().app}.` : "Sign in through the CLI on the bridge machine, then check the connection again.";
      byId("login-card-title").textContent = PROVIDER_NAMES[provider];
      byId("login-badge").textContent = provider === "claude" ? "BROWSER" : "DEVICE CODE";
      byId("login-device").hidden = !inApp;
      byId("login-manual").hidden = inApp;
      byId("login-command").textContent = PROVIDER_LOGIN_COMMANDS[provider];
      byId("login-manual-hint").textContent = PROVIDER_LOGIN_HINTS[provider];
      if (inApp) deps.loginPanel.open();
      else deps.loginPanel.close();
    } else deps.loginPanel.close();
    byId("tab-chat").setAttribute("aria-pressed", String(name === null));
    byId("tab-changes").setAttribute("aria-pressed", "false");
    byId("tab-history").setAttribute("aria-pressed", String(name === "session"));
    if (name === "templates") void deps.loadTemplateSkills();
  }

  function openMenu(): void {
    deps.renderProviderMenu();
    els.menu.hidden = false;
    els.menuButton.setAttribute("aria-expanded", "true");
  }

  function closeMenu(): void {
    els.menu.hidden = true;
    els.menuButton.setAttribute("aria-expanded", "false");
  }
  return { showPanel, openPanel, backFromPanel, openFloating, closeFloating, openMenu, closeMenu };
}
