// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import type { ProviderConnectionStatus, ProviderId } from "../protocol/src/index.js";
import { setupPanels, type PanelName } from "../addin/taskpane/ui/panels.js";

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
    loginPanel,
    elements: {
      panels: Object.fromEntries(names.map(name => [name, element(`panel-${name}`)])) as Record<PanelName, HTMLElement>,
      menu: element("settings-menu"), menuButton: element<HTMLButtonElement>("menu-button"),
    },
  });
  return { api, loginPanel };
}

describe("Setting up a provider that is not installed", () => {
  it.each(["claude", "copilot", "codex", "opencode"] as const)("shows the install command for %s and starts no sign-in", provider => {
    const { api, loginPanel } = panels(provider, "not_installed");
    api.showPanel("login");
    expect(loginPanel.open).not.toHaveBeenCalled();
    expect(element("login-heading").textContent).toMatch(/^Set up /);
    expect(element("login-device").hidden).toBe(true);
    expect(element("login-manual").hidden).toBe(false);
    expect(element("login-command").textContent).toBe(`office-partner providers install ${provider}`);
    expect(element("login-recheck").textContent).toBe("Check again");
  });

  it.each(["claude", "copilot", "codex"] as const)("still signs in to an installed %s from the pane", provider => {
    const { api, loginPanel } = panels(provider, "login_required");
    api.showPanel("login");
    expect(loginPanel.open).toHaveBeenCalledOnce();
    expect(element("login-heading").textContent).toMatch(/^Sign in to /);
    expect(element("login-device").hidden).toBe(false);
    expect(element("login-manual").hidden).toBe(true);
  });

  it("keeps the terminal login command for OpenCode once it is installed", () => {
    const { api, loginPanel } = panels("opencode", "login_required");
    api.showPanel("login");
    expect(loginPanel.open).not.toHaveBeenCalled();
    expect(element("login-command").textContent).toBe("office-partner providers login opencode");
    expect(element("login-recheck").textContent).toBe("Check sign-in status");
  });
});
