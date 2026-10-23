import { supportsNativeWebSearch, type ProviderId, type ToolCapabilities } from "../../../protocol/src/index.js";

const STORAGE_KEY = "excelPartner.toolCapabilities";

export function setupToolSettings(provider: () => ProviderId = () => "claude"): { value(): ToolCapabilities; refreshProvider(): void } {
  const button = document.getElementById("tools-button") as HTMLButtonElement;
  const menu = document.getElementById("tools-menu")!;
  const search = document.getElementById("tool-web-search") as HTMLInputElement;
  const browser = document.getElementById("tool-playwright") as HTMLInputElement;
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "{}");
    search.checked = saved?.webSearch === true;
    browser.checked = saved?.playwright === true;
  } catch { /* invalid settings default to off */ }
  let preferredWebSearch = search.checked;
  const value = () => ({ webSearch: search.checked && supportsNativeWebSearch(provider()), playwright: browser.checked });
  const refresh = () => {
    const supported = supportsNativeWebSearch(provider());
    search.disabled = !supported;
    search.checked = supported && preferredWebSearch;
    document.getElementById("tool-web-search-description")!.textContent = supported
      ? "Use the selected AI provider’s built-in web search. No separate search service is added."
      : "Unavailable for OpenCode: its search requires an external search backend.";
    const count = Number(value().webSearch) + Number(browser.checked);
    button.querySelector(".picker-label")!.textContent = count ? `Tools (${count})` : "Tools";
    button.setAttribute("aria-label", count ? `Tools (${count} enabled)` : "Tools");
    button.dataset.active = String(count > 0);
    button.title = `Choose tools for the next message (${count} enabled)`;
  };
  const close = () => { menu.hidden = true; button.setAttribute("aria-expanded", "false"); };
  const position = () => {
    const bounds = menu.getBoundingClientRect();
    const parent = menu.offsetParent as HTMLElement;
    const left = Math.max(8, Math.min(bounds.left, window.innerWidth - bounds.width - 8));
    menu.style.left = `${left - parent.getBoundingClientRect().left}px`;
    menu.style.right = "auto";
  };
  button.addEventListener("click", () => {
    if (!menu.hidden) return close();
    document.dispatchEvent(new CustomEvent("composer-picker-open", { detail: button.id }));
    menu.hidden = false;
    position();
    button.setAttribute("aria-expanded", "true");
    (search.disabled ? browser : search).focus();
  });
  for (const input of [search, browser]) input.addEventListener("change", () => {
    if (input === search) preferredWebSearch = search.checked;
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ webSearch: preferredWebSearch, playwright: browser.checked }));
    refresh();
  });
  menu.addEventListener("keydown", event => {
    if (event.key === "Escape") { event.preventDefault(); close(); button.focus(); }
  });
  document.addEventListener("pointerdown", event => {
    if (!button.contains(event.target as Node) && !menu.contains(event.target as Node)) close();
  });
  document.addEventListener("composer-picker-open", event => {
    if ((event as CustomEvent<string>).detail !== button.id) close();
  });
  window.addEventListener("resize", () => { if (!menu.hidden) position(); });
  refresh();
  return { value, refreshProvider: refresh };
}
