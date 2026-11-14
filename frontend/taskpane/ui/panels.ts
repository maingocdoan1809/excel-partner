import type { ProviderCapabilities, ProviderConnectionStatus, ProviderId, ProviderLoginKind } from "../../../protocol/src/index.js";
import { activeHost } from "../host/index.js";
import { byId } from "./dom.js";
import { PROVIDER_NAMES } from "./provider-picker.js";

export type PanelName = "settings" | "skills" | "login" | "changes" | "history" | "references" | "templates" | "session" | "preview" | "browser" | "testing";

const LOGIN_BADGES: Record<ProviderLoginKind, string> = { device: "DEVICE CODE", browser: "BROWSER", account: "ACCOUNT" };

export interface PanelsDeps {
  getProvider(): ProviderId;
  getProviderStatus(provider: ProviderId): ProviderConnectionStatus;
  getCapabilities(provider: ProviderId): ProviderCapabilities | undefined;
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
  loginInstall: { reset(): void };
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
      const name = PROVIDER_NAMES[provider];
      // A CLI that is not on this computer cannot sign in: offer to install it instead of starting a login.
      const notInstalled = deps.getProviderStatus(provider) === "not_installed";
      byId("login-heading").textContent = notInstalled ? `Set up ${name}` : `Sign in to ${name}`;
      byId("login-description").textContent = notInstalled ? `${name} is not installed on this computer.`
        : `Finish signing in through your browser, then return to ${activeHost().app}.`;
      byId("login-card-title").textContent = name;
      byId("login-badge").textContent = LOGIN_BADGES[deps.getCapabilities(provider)?.loginKind ?? "device"];
      byId("login-device").hidden = notInstalled;
      byId("login-manual").hidden = !notInstalled;
      byId("login-manual-hint").textContent = `Installs a private copy of ${name} for Office Partner. This can take a few minutes.`;
      const install = byId<HTMLButtonElement>("login-install");
      install.hidden = !notInstalled;
      install.textContent = `Install ${name}`;
      deps.loginInstall.reset();
      if (notInstalled) deps.loginPanel.close();
      else deps.loginPanel.open();
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
