import type { SkillTextFile } from "../bridge-client.js";
import { recommendedPromptFor } from "../../../providers/src/host-prompts.js";
import { activeHost } from "../host/index.js";

interface SettingsOptions {
  loadInstructions(): Promise<SkillTextFile>;
  saveInstructions(content: string, revision: string): Promise<SkillTextFile>;
  historyCount(): number;
  clearHistory(): void;
  historyClearButton: HTMLButtonElement;
}

/** Pane preferences are local; instructions use the existing revisioned bridge API. */
export function setupSettings(options: SettingsOptions): { open(preserveScroll?: boolean): Promise<void> } {
  const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const input = get<HTMLTextAreaElement>("settings-instructions");
  const save = get<HTMLButtonElement>("settings-save-instructions");
  const status = get("settings-instructions-status");
  const clear = get<HTMLButtonElement>("settings-clear-history");
  const size = get<HTMLButtonElement>("text-size-reset");
  let loaded: SkillTextFile | undefined;
  let busy = false;
  const dirty = () => loaded !== undefined && input.value !== loaded.content;
  const updateSave = () => {
    save.disabled = busy || !dirty();
    recommended.disabled = busy || !loaded;
    input.setAttribute("aria-busy", String(busy));
  };
  const setStatus = (message: string, state = "idle") => { status.textContent = message; status.dataset.state = state; };
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
  input.addEventListener("input", () => { setStatus(dirty() ? "Unsaved changes" : ""); updateSave(); });
  const recommended = get<HTMLButtonElement>("settings-recommended-instructions");
  recommended.addEventListener("click", () => {
    if (!loaded || busy) return;
    input.value = recommendedPromptFor(activeHost().id);
    setStatus("Recommended instructions loaded. Save to apply them.");
    updateSave();
  });
  save.addEventListener("click", async () => {
    if (!loaded || busy || !dirty()) return;
    busy = true; input.disabled = true; save.textContent = "Saving…"; updateSave(); setStatus("Saving…");
    try {
      loaded = await options.saveInstructions(input.value, loaded.revision);
      input.value = loaded.content;
      setStatus("Saved. Applies to your next message; use a new chat to replace earlier instructions.", "saved");
    } catch (error) { setStatus(error instanceof Error ? error.message : "Could not save instructions.", "error"); }
    finally { busy = false; input.disabled = false; save.textContent = "Save instructions"; updateSave(); }
  });
  updateSave();
  return { async open(preserveScroll = false) {
    const body = document.querySelector<HTMLElement>(".settings-body")!;
    if (!preserveScroll) { body.scrollTop = 0; get("settings-heading").focus({ preventScroll: true }); }
    history();
    if (busy || dirty() || (preserveScroll && loaded)) return;
    busy = true; input.disabled = true; updateSave(); setStatus("Loading…");
    try { loaded = await options.loadInstructions(); input.value = loaded.content; setStatus(""); }
    catch (error) { loaded = undefined; setStatus(error instanceof Error ? error.message : "Could not load instructions.", "error"); }
    finally { busy = false; input.disabled = !loaded; updateSave(); }
  } };
}
