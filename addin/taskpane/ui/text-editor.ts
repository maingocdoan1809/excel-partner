import type { SkillTextFile } from "../bridge-client.js";

interface EditorOptions {
  title: string;
  detail: string;
  files: string[];
  load(file: string): Promise<SkillTextFile>;
  save(document: SkillTextFile): Promise<SkillTextFile>;
  reset?(revision: string): Promise<SkillTextFile>;
  saved?(): void;
}

/** Shared local text editor; never renders the edited contents as HTML. */
export function createTextEditor(): { open(options: EditorOptions): Promise<void> } {
  const dialog = document.createElement("dialog");
  dialog.className = "reference-dialog text-editor";
  dialog.setAttribute("aria-label", "Text editor");
  dialog.innerHTML = '<form><h2></h2><p class="template-help"></p><label>File<select aria-label="File"></select></label><textarea aria-label="File content" spellcheck="false" rows="16"></textarea><p class="template-help editor-status" role="status" aria-live="polite"></p><div class="reference-actions"><button type="button" class="secondary editor-reset">Restore default</button><button type="button" class="secondary editor-cancel">Close</button><button type="submit">Save</button></div></form>';
  document.body.append(dialog);
  const form = dialog.querySelector("form")!;
  const heading = dialog.querySelector("h2")!;
  const detail = dialog.querySelector("p")!;
  const select = dialog.querySelector("select")!;
  const text = dialog.querySelector("textarea")!;
  const status = dialog.querySelector<HTMLElement>(".editor-status")!;
  const save = dialog.querySelector<HTMLButtonElement>('button[type="submit"]')!;
  const close = dialog.querySelector<HTMLButtonElement>(".editor-cancel")!;
  const reset = dialog.querySelector<HTMLButtonElement>(".editor-reset")!;
  let options: EditorOptions;
  let loaded: SkillTextFile | null = null;
  let busy = false;
  let discard = false;
  const dirty = () => loaded !== null && text.value !== loaded.content;
  const errorText = (error: unknown) => error instanceof Error ? error.message : "Could not update this file.";

  function setBusy(value: boolean): void {
    busy = value;
    select.disabled = text.disabled = save.disabled = reset.disabled = value;
    close.disabled = value;
  }
  function show(value: SkillTextFile): void {
    loaded = value;
    text.value = value.content;
    select.value = value.file;
    discard = false;
    close.textContent = "Close";
  }
  async function load(file: string): Promise<void> {
    setBusy(true);
    status.textContent = "Loading…";
    try { show(await options.load(file)); status.textContent = ""; }
    catch (error) { status.textContent = errorText(error); if (loaded) select.value = loaded.file; }
    finally { setBusy(false); save.disabled = loaded === null; }
  }
  function cancel(): void {
    if (busy) return;
    if (dirty() && !discard) {
      discard = true;
      close.textContent = "Discard edits";
      status.textContent = "You have unsaved changes. Save them or choose Discard edits to close.";
      return;
    }
    dialog.close();
  }
  close.addEventListener("click", cancel);
  dialog.addEventListener("cancel", event => { event.preventDefault(); cancel(); });
  text.addEventListener("input", () => { discard = false; close.textContent = "Close"; status.textContent = "Unsaved changes"; });
  select.addEventListener("change", () => {
    if (dirty()) { select.value = loaded!.file; status.textContent = "Save this file before switching to another file."; return; }
    void load(select.value);
  });
  form.addEventListener("submit", async event => {
    event.preventDefault();
    if (busy || !loaded) return;
    setBusy(true);
    status.textContent = "Saving…";
    try { show(await options.save({ ...loaded, content: text.value })); status.textContent = "Saved."; options.saved?.(); }
    catch (error) { status.textContent = errorText(error); }
    finally { setBusy(false); }
  });
  reset.addEventListener("click", async () => {
    if (busy || !loaded || !options.reset) return;
    if (dirty()) { status.textContent = "Save or discard your edits before restoring the default."; return; }
    setBusy(true);
    try { show(await options.reset(loaded.revision)); status.textContent = "Default restored."; options.saved?.(); }
    catch (error) { status.textContent = errorText(error); }
    finally { setBusy(false); }
  });
  return { async open(next) {
    if (dialog.open) return;
    options = next;
    loaded = null;
    text.value = "";
    heading.textContent = next.title;
    detail.textContent = next.detail;
    select.replaceChildren(...next.files.map(file => {
      const option = window.document.createElement("option");
      option.value = option.textContent = file;
      return option;
    }));
    select.parentElement!.hidden = next.files.length <= 1;
    reset.hidden = !next.reset;
    dialog.showModal();
    if (next.files[0]) await load(next.files[0]);
    else { status.textContent = "No editable text files in this skill."; save.disabled = true; }
  } };
}
