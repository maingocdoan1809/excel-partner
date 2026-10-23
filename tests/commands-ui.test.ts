// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createSlashMenu } from "../addin/taskpane/ui/slash-menu.js";
import { normalizeCommands } from "../providers/src/commands.js";
import type { ProviderCommand } from "../protocol/src/commands.js";

beforeEach(() => {
  document.body.innerHTML = '<textarea></textarea><div id="menu"></div>';
  HTMLElement.prototype.scrollIntoView = vi.fn();
  vi.stubGlobal("fetch", vi.fn(async () => ({ ok: true, json: async () => ({ skills: [{ name: "skills", description: "Bridge skill" }] }) })));
});
afterEach(() => { vi.unstubAllGlobals(); document.body.replaceChildren(); });

describe("provider slash menu", () => {
  it("exposes all commands and keeps a colliding bridge skill addressable", async () => {
    const input = document.querySelector("textarea")!;
    const menu = document.querySelector<HTMLElement>("#menu")!;
    const catalog = normalizeCommands(Array.from({ length: 35 }, (_, index) => ({ name: index === 0 ? "skills" : `command-${index}`, description: "Native" })));
    const slash = createSlashMenu({ input, menu, onBuiltin() {}, getProvider: () => "copilot", loadCommands: async () => catalog });
    expect(await slash.resolve("/skills list")).toMatchObject({ kind: "provider" });
    expect(await slash.resolve("/skill:skills task")).toMatchObject({ kind: "skill", name: "skill:skills" });
    input.value = "/"; slash.update();
    await vi.waitFor(() => expect(menu.querySelectorAll(".slash-item").length).toBeGreaterThan(30));
    expect(menu.textContent).toContain("/command-34");
  });
  it("does not select unsupported terminal commands as runnable prompts", async () => {
    const input = document.querySelector("textarea")!;
    const menu = document.querySelector<HTMLElement>("#menu")!;
    const unavailable = vi.fn(), select = vi.fn();
    const slash = createSlashMenu({ input, menu, onBuiltin() {}, onSelect: select, onUnavailable: unavailable,
      loadCommands: async () => normalizeCommands([{ name: "vim", description: "Vim", execution: "unsupported", reason: "Terminal only" }]) });
    await slash.resolve("/vim"); input.value = "/vim"; slash.update();
    await vi.waitFor(() => expect(menu.querySelector('[aria-disabled="true"]')).not.toBeNull());
    slash.handleKey(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));
    expect(unavailable).toHaveBeenCalledWith("Terminal only"); expect(select).not.toHaveBeenCalled();
  });
  it("discards a late catalog from the previous provider", async () => {
    const input = document.querySelector("textarea")!;
    const menu = document.querySelector<HTMLElement>("#menu")!;
    let provider = "old";
    let finishOld!: (commands: ProviderCommand[]) => void;
    const old = new Promise<ProviderCommand[]>(resolve => { finishOld = resolve; });
    const slash = createSlashMenu({ input, menu, onBuiltin() {}, getProvider: () => provider,
      loadCommands: () => provider === "old" ? old : Promise.resolve(normalizeCommands([{ name: "current", description: "Current" }])) });
    input.value = "/"; slash.update(); provider = "new"; slash.update();
    await vi.waitFor(() => expect(menu.textContent).toContain("/current"));
    finishOld(normalizeCommands([{ name: "stale", description: "Stale" }]));
    await old; await Promise.resolve(); await Promise.resolve();
    expect(slash.commands.map(command => command.name)).not.toContain("stale");
  });
});
