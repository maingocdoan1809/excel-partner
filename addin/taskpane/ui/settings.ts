import type { SkillTextFile } from "../bridge-client.js";
import { RECOMMENDED_SYSTEM_PROMPT } from "../../../providers/src/prompt.js";

interface SettingsOptions {
  loadInstructions(): Promise<SkillTextFile>;
  saveInstructions(content: string, revision: string): Promise<SkillTextFile>;
  historyCount(): number;
  clearHistory(): void;
  historyClearButton: HTMLButtonElement;
}

/** Pane preferences are local; instructions use the existing revisioned bridge API. */
export function setupSettings(options: SettingsOptions): { open(): Promise<void> } {
  const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const input = get<HTMLTextAreaElement>("settings-instructions");
  const save = get<HTMLButtonElement>("settings-save-instructions");
  const status = get("settings-instructions-status");
  const clear = get<HTMLButtonElement>("settings-clear-history");
  const size = get<HTMLButtonElement>("text-size-reset");
  let loaded: SkillTextFile | undefined;
  let busy = false;
  const dirty = () => loaded !== undefined && input.value !== loaded.content;
  const updateSave = () => { save.disabled = busy || !dirty(); };
  const history = () => {
    const count = options.historyCount();
    get("settings-history-count").textContent = `${count} saved chat${count === 1 ? "" : "s"}`;
    clear.disabled = count === 0;
    clear.textContent = options.historyClearButton.textContent;
  };
  const observer = new MutationObserver(history);
  observer.observe(options.historyClearButton, { childList: true, subtree: true });
  window.addEventListener("pagehide", () => observer.disconnect(), { once: true });
  clear.addEventListener("click", () => { options.clearHistory(); history(); });
  let scale = 100;
  try { const stored = Number(localStorage.getItem("excelPartner.textSize")); if (stored >= 80 && stored <= 150) scale = stored; } catch {}
  const applySize = () => {
    document.documentElement.style.setProperty("--chat-font-size", `${13 * scale / 100}px`);
    size.textContent = `${scale}%`;
    get<HTMLButtonElement>("text-size-decrease").disabled = scale <= 80;
    get<HTMLButtonElement>("text-size-increase").disabled = scale >= 150;
    try { localStorage.setItem("excelPartner.textSize", String(scale)); } catch {}
  };
  get("text-size-decrease").addEventListener("click", () => { scale = Math.max(80, scale - 10); applySize(); });
  get("text-size-increase").addEventListener("click", () => { scale = Math.min(150, scale + 10); applySize(); });
  size.addEventListener("click", () => { scale = 100; applySize(); });
  applySize();
  input.addEventListener("input", () => { status.textContent = "Unsaved changes"; updateSave(); });
  const recommended = get<HTMLButtonElement>("settings-recommended-instructions");
  recommended.addEventListener("click", () => {
    if (!loaded || busy) return;
    input.value = RECOMMENDED_SYSTEM_PROMPT;
    status.textContent = "Recommended instructions loaded. Press Save to apply them.";
    updateSave();
  });
  save.addEventListener("click", async () => {
    if (!loaded || busy || !dirty()) return;
    busy = true; input.disabled = true; updateSave(); status.textContent = "Saving…";
    try {
      loaded = await options.saveInstructions(input.value, loaded.revision);
      input.value = loaded.content;
      status.textContent = "Saved. Applies to your next message; use a new chat to replace earlier instructions.";
    } catch (error) { status.textContent = error instanceof Error ? error.message : "Could not save instructions."; }
    finally { busy = false; input.disabled = false; updateSave(); }
  });
  updateSave();
  return { async open() {
    const body = document.querySelector<HTMLElement>(".settings-body")!;
    body.scrollTop = 0;
    get("settings-heading").focus({ preventScroll: true });
    history();
    if (busy || dirty()) return;
    busy = true; input.disabled = true; updateSave(); status.textContent = "Loading…";
    try { loaded = await options.loadInstructions(); input.value = loaded.content; status.textContent = ""; }
    catch (error) { loaded = undefined; status.textContent = error instanceof Error ? error.message : "Could not load instructions."; }
    finally { busy = false; input.disabled = !loaded; updateSave(); }
  } };
}
