import type { ToolCapabilities } from "../../../protocol/src/index.js";

const STORAGE_KEY = "excelPartner.toolCapabilities";

export function setupToolSettings(webSearchSupported: () => boolean): { value(): ToolCapabilities; refreshProvider(): void; setBridgeDefault(lazy: boolean): void } {
  const button = document.getElementById("tools-button") as HTMLButtonElement;
  const search = document.getElementById("tool-web-search") as HTMLInputElement;
  const browser = document.getElementById("tool-playwright") as HTMLInputElement;
  const lazy = document.getElementById("tool-lazy-schemas") as HTMLInputElement;
  // Until the person chooses, the bridge's own setting decides and the pane says nothing about it.
  let chosenLazy: boolean | undefined;
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "{}");
    search.checked = saved?.webSearch === true;
    browser.checked = saved?.playwright === true;
    if (typeof saved?.lazySchemas === "boolean") chosenLazy = saved.lazySchemas;
  } catch { /* invalid settings default to off */ }
  lazy.checked = chosenLazy === true;
  let preferredWebSearch = search.checked;
  const value = (): ToolCapabilities => ({ webSearch: search.checked && webSearchSupported(), playwright: browser.checked, ...(chosenLazy === undefined ? {} : { lazySchemas: chosenLazy }) });
  const refresh = () => {
    const supported = webSearchSupported();
    search.disabled = !supported;
    search.checked = supported && preferredWebSearch;
    document.getElementById("tool-web-search-description")!.textContent = supported
      ? "Use the selected AI provider’s built-in web search. No separate search service is added."
      : "Unavailable for this provider: it has no built-in web search.";
    const count = Number(value().webSearch) + Number(browser.checked);
    button.querySelector(".picker-label")!.textContent = count ? `Tools (${count})` : "Tools";
    button.setAttribute("aria-label", count ? `Tools (${count} enabled)` : "Tools");
    button.dataset.active = String(count > 0);
    button.title = `Choose tools for the next message (${count} enabled)`;
  };
  for (const input of [search, browser, lazy]) input.addEventListener("change", () => {
    if (input === search) preferredWebSearch = search.checked;
    if (input === lazy) chosenLazy = lazy.checked;
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ webSearch: preferredWebSearch, playwright: browser.checked, ...(chosenLazy === undefined ? {} : { lazySchemas: chosenLazy }) })); } catch { /* the choice still applies to this pane */ }
    refresh();
  });
  refresh();
  return { value, refreshProvider: refresh, setBridgeDefault: (isLazy) => { if (chosenLazy === undefined) lazy.checked = isLazy; } };
}
