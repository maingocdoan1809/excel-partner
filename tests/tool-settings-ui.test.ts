// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from "vitest";
import { setupToolSettings } from "../frontend/taskpane/ui/tool-settings.js";

const STORAGE_KEY = "excelPartner.toolCapabilities";

function mount(stored?: object) {
  localStorage.clear();
  if (stored) localStorage.setItem(STORAGE_KEY, JSON.stringify(stored));
  document.body.innerHTML = `
    <button id="tools-button"><span class="picker-label"></span></button>
    <input id="tool-web-search" type="checkbox" /><p id="tool-web-search-description"></p>
    <input id="tool-playwright" type="checkbox" />
    <input id="tool-lazy-schemas" type="checkbox" />`;
  return {
    settings: setupToolSettings(() => true),
    lazy: document.getElementById("tool-lazy-schemas") as HTMLInputElement,
  };
}

describe("compact tool list setting", () => {
  beforeEach(() => localStorage.clear());

  it("says nothing about it until the person chooses, and shows what the bridge does", () => {
    const { settings, lazy } = mount();
    expect(settings.value()).toEqual({ webSearch: false, playwright: false });
    settings.setBridgeDefault(true);
    expect(lazy.checked).toBe(true);
    expect(settings.value()).not.toHaveProperty("lazySchemas");
  });

  it("sends the person's choice, keeps it, and no longer follows the bridge", () => {
    const first = mount();
    first.settings.setBridgeDefault(true);
    first.lazy.checked = false;
    first.lazy.dispatchEvent(new Event("change"));
    expect(first.settings.value()).toMatchObject({ lazySchemas: false });

    const again = mount(JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "{}"));
    again.settings.setBridgeDefault(true);
    expect(again.lazy.checked).toBe(false);
    expect(again.settings.value()).toMatchObject({ lazySchemas: false });
  });

  it("is not counted as a tool on the Tools button", () => {
    const { lazy } = mount();
    lazy.checked = true;
    lazy.dispatchEvent(new Event("change"));
    expect(document.querySelector(".picker-label")!.textContent).toBe("Tools");
  });
});
