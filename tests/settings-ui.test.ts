// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import { setupSettings } from "../addin/taskpane/ui/settings.js";
import { renderToolCard, formatToolValue } from "../addin/taskpane/ui/tool-view.js";

beforeEach(() => {
  const html = fs.readFileSync("addin/taskpane/index.html", "utf8");
  document.body.innerHTML = html.split("<body>")[1]!.split("</body>")[0]!;
  localStorage.clear();
});
afterEach(() => { window.dispatchEvent(new Event("pagehide")); document.body.replaceChildren(); vi.restoreAllMocks(); });
const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const options = () => ({ loadInstructions: vi.fn(async () => ({ file: "System prompt", content: "Original", revision: "rev-1" })),
  saveInstructions: vi.fn(async (content: string, revision: string) => ({ file: "System prompt", content, revision: `${revision}-next` })),
  historyCount: () => 3, clearHistory: vi.fn(), historyClearButton: element<HTMLButtonElement>("history-clear") });

describe("Settings and tool details", () => {
  it("loads and saves instructions using the current bridge revision", async () => {
    const api = options(); const settings = setupSettings(api); await settings.open();
    const input = element<HTMLTextAreaElement>("settings-instructions");
    expect(input.value).toBe("Original");
    expect(element<HTMLButtonElement>("settings-save-instructions").disabled).toBe(true);
    input.value = "Use Vietnamese"; input.dispatchEvent(new Event("input"));
    element("settings-save-instructions").click();
    await vi.waitFor(() => expect(element("settings-instructions-status").textContent).toContain("Saved."));
    expect(api.saveInstructions).toHaveBeenCalledWith("Use Vietnamese", "rev-1");
  });
  it("preserves unsaved instructions when reopened and reports save conflicts", async () => {
    const api = options(); api.saveInstructions.mockRejectedValue(new Error("Instructions were changed by another editor"));
    const settings = setupSettings(api); await settings.open();
    const input = element<HTMLTextAreaElement>("settings-instructions");
    input.value = "Unsaved text"; input.dispatchEvent(new Event("input"));
    await settings.open(); expect(api.loadInstructions).toHaveBeenCalledTimes(1);
    element("settings-save-instructions").click();
    await vi.waitFor(() => expect(element("settings-instructions-status").textContent).toContain("another editor"));
    expect(input.value).toBe("Unsaved text");
    expect(element<HTMLButtonElement>("settings-save-instructions").disabled).toBe(false);
  });
  it("changes and resets persisted text size and delegates history deletion", async () => {
    const api = options(); await setupSettings(api).open();
    element("text-size-increase").click();
    expect(localStorage.getItem("excelPartner.textSize")).toBe("110");
    expect(document.documentElement.style.getPropertyValue("--chat-font-size")).toBe("14.3px");
    element("text-size-reset").click(); expect(localStorage.getItem("excelPartner.textSize")).toBe("100");
    expect(element("settings-history-count").textContent).toBe("3 saved chats");
    element("settings-clear-history").click(); expect(api.clearHistory).toHaveBeenCalledOnce();
  });
  it("renders actual tool data safely in expandable details and bounds large values", () => {
    const result = formatToolValue({ ok: true, data: { sheet: "<img src=x onerror=alert(1)>" } });
    const card = renderToolCard({ name: "read_range", args: "A1", input: '{"address":"A1"}', result, error: null, state: "ok" });
    document.body.append(card);
    expect(card.open).toBe(false);
    expect(card.querySelector("summary")?.textContent).toContain("Reading cells");
    expect(card.querySelector("img")).toBeNull();
    expect(card.textContent).toContain("<img src=x onerror=alert(1)>");
    expect(card.querySelector(".json-key")).not.toBeNull();
    expect(formatToolValue({ type: "image", source: { data: "PRIVATE_IMAGE" } })).not.toContain("PRIVATE_IMAGE");
    expect(formatToolValue({ values: "x".repeat(20000) }).length).toBeLessThan(16050);
  });
});
