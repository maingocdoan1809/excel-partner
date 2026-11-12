// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import type { ProviderConnectionStatus, ProviderId } from "../protocol/src/index.js";
import { setupPanels, type PanelName } from "../addin/taskpane/ui/panels.js";
import { setupProviderInstall } from "../addin/taskpane/ui/provider-install.js";

beforeEach(() => {
  const html = fs.readFileSync("addin/taskpane/index.html", "utf8");
  document.body.innerHTML = html.split("<body>")[1]!.split("</body>")[0]!;
});
afterEach(() => { document.body.replaceChildren(); vi.restoreAllMocks(); });
const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

function panels(provider: ProviderId, status: ProviderConnectionStatus) {
  const names: PanelName[] = ["settings", "skills", "login", "changes", "history", "references", "templates", "session", "preview", "browser", "testing"];
  const loginPanel = { open: vi.fn(), close: vi.fn() };
  const api = setupPanels({
    getProvider: () => provider,
    getProviderStatus: () => status,
    renderHistory: vi.fn(), renderChanges: vi.fn(), renderProviderMenu: vi.fn(), openSettings: vi.fn(), closePreview: vi.fn(),
    refreshReferences: vi.fn(async () => {}), refreshSkills: vi.fn(async () => {}), refreshTesting: vi.fn(async () => {}), loadTemplateSkills: vi.fn(async () => {}),
    loginPanel, loginInstall: { reset: vi.fn() },
    elements: {
      panels: Object.fromEntries(names.map(name => [name, element(`panel-${name}`)])) as Record<PanelName, HTMLElement>,
      menu: element("settings-menu"), menuButton: element<HTMLButtonElement>("menu-button"),
    },
  });
  return { api, loginPanel };
}

describe("Setting up a provider that is not installed", () => {
  it.each(["claude", "copilot", "codex", "opencode"] as const)("offers only an Install button for %s and starts no sign-in", provider => {
    const { api, loginPanel } = panels(provider, "not_installed");
    api.showPanel("login");
    expect(loginPanel.open).not.toHaveBeenCalled();
    expect(element("login-heading").textContent).toMatch(/^Set up /);
    expect(element("login-device").hidden).toBe(true);
    expect(element("login-manual").hidden).toBe(false);
    expect(element("login-command").hidden).toBe(true);
    expect(element("login-recheck").hidden).toBe(true);
    expect(element<HTMLButtonElement>("login-install").hidden).toBe(false);
    expect(element("login-install").textContent).toMatch(/^Install /);
  });

  it.each(["claude", "copilot", "codex"] as const)("still signs in to an installed %s from the pane", provider => {
    const { api, loginPanel } = panels(provider, "login_required");
    api.showPanel("login");
    expect(loginPanel.open).toHaveBeenCalledOnce();
    expect(element("login-heading").textContent).toMatch(/^Sign in to /);
    expect(element("login-device").hidden).toBe(false);
    expect(element("login-manual").hidden).toBe(true);
    expect(element<HTMLButtonElement>("login-install").hidden).toBe(true);
  });

  it("keeps the terminal login command for OpenCode once it is installed", () => {
    const { api, loginPanel } = panels("opencode", "login_required");
    api.showPanel("login");
    expect(loginPanel.open).not.toHaveBeenCalled();
    expect(element("login-command").textContent).toBe("office-partner providers login opencode");
    expect(element("login-command").hidden).toBe(false);
    expect(element("login-recheck").hidden).toBe(false);
  });
});

describe("Installing a provider from the set-up panel", () => {
  function install(systemAction: (action: string, provider?: ProviderId) => Promise<void>, busy = false) {
    const onInstalled = vi.fn();
    setupProviderInstall({ client: () => ({ systemAction }), provider: () => "codex", isBusy: () => busy, onInstalled });
    return { onInstalled, click: () => element("login-install").click() };
  }

  it("runs the same install action as Settings > System & recovery", async () => {
    const systemAction = vi.fn(async () => {});
    const { onInstalled, click } = install(systemAction);
    click();
    expect(element("login-install-status").textContent).toContain("Installing");
    expect(element<HTMLButtonElement>("login-install").disabled).toBe(true);
    await vi.waitFor(() => expect(onInstalled).toHaveBeenCalledWith("codex", false));
    expect(systemAction).toHaveBeenCalledWith("install", "codex");
    expect(element<HTMLButtonElement>("login-install").disabled).toBe(false);
  });

  it("treats an installed CLI that needs sign-in as installed", async () => {
    const { onInstalled, click } = install(vi.fn(async () => { throw Object.assign(new Error("Private CLI installed. Sign in."), { code: "login_required" }); }));
    click();
    await vi.waitFor(() => expect(onInstalled).toHaveBeenCalledWith("codex", true));
  });

  it("keeps the button and shows the reason when the install fails", async () => {
    const { onInstalled, click } = install(vi.fn(async () => { throw new Error("The operation did not finish."); }));
    click();
    await vi.waitFor(() => expect(element("login-install-status").textContent).toBe("The operation did not finish."));
    expect(element("login-install-status").dataset.state).toBe("error");
    expect(element<HTMLButtonElement>("login-install").disabled).toBe(false);
    expect(onInstalled).not.toHaveBeenCalled();
  });

  it("does not start while a response is running", () => {
    const systemAction = vi.fn(async () => {});
    const { click } = install(systemAction, true);
    click();
    expect(systemAction).not.toHaveBeenCalled();
    expect(element("login-install-status").textContent).toContain("current response");
  });
});
