import type { BridgeClient, SavedReference } from "../bridge-client.js";

export function setupSavedReferences(client: () => BridgeClient): { refresh(): Promise<void> } {
  const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const list = get<HTMLElement>("reference-list");
  const status = get<HTMLElement>("reference-status");
  const dialog = get<HTMLDialogElement>("reference-dialog");
  const form = get<HTMLFormElement>("reference-form");
  const name = get<HTMLInputElement>("reference-name");
  const file = get<HTMLInputElement>("reference-path");
  const description = get<HTMLTextAreaElement>("reference-description");
  const message = get<HTMLElement>("reference-dialog-status");
  const save = get<HTMLButtonElement>("reference-save");
  const cancel = get<HTMLButtonElement>("reference-cancel");
  let editing: string | undefined;
  let originalPath = "";
  let busy = false;
  let request = 0;
  const errorText = (error: unknown) => error instanceof Error ? error.message : "Could not update reference files.";

  function setBusy(value: boolean): void {
    busy = value;
    file.disabled = save.disabled = cancel.disabled = value;
  }

  function open(entry?: SavedReference): void {
    editing = entry?.id;
    originalPath = entry?.path ?? "";
    name.value = entry?.name ?? "";
    file.value = entry?.path ?? "";
    description.value = entry?.description ?? "";
    message.textContent = entry && !entry.available ? "File not found. Enter its new path." : "";
    dialog.showModal();
    (entry ? name : file).focus();
  }

  async function refresh(): Promise<void> {
    const current = ++request;
    list.textContent = "Loading reference files…";
    try {
      const entries = await client().listSavedReferences();
      if (current !== request) return;
      list.replaceChildren();
      if (!entries.length) list.textContent = "No reference files registered yet.";
      for (const entry of entries) {
        const row = document.createElement("div");
        row.className = "reference-row";
        const title = document.createElement("strong");
        title.textContent = entry.name;
        const location = document.createElement("p");
        location.textContent = entry.path;
        const detail = document.createElement("p");
        detail.textContent = entry.description;
        row.append(title, location, detail);
        if (!entry.available) {
          const missing = document.createElement("p");
          missing.className = "reference-missing";
          missing.textContent = "File not found — edit the reference to enter its new path.";
          row.append(missing);
        }
        const actions = document.createElement("div");
        actions.className = "reference-actions";
        const edit = document.createElement("button");
        edit.type = "button";
        edit.className = "secondary";
        edit.textContent = "Edit";
        edit.addEventListener("click", () => open(entry));
        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "secondary";
        remove.textContent = "Remove link";
        remove.addEventListener("click", async () => {
          if (remove.disabled) return;
          remove.disabled = true;
          try {
            await client().removeReference(entry.id);
            status.textContent = "Reference link removed.";
            await refresh();
          } catch (error) { status.textContent = errorText(error); }
          finally { remove.disabled = false; }
        });
        actions.append(edit, remove);
        row.append(actions);
        list.append(row);
      }
    } catch (error) { if (current === request) list.textContent = errorText(error); }
  }

  get<HTMLButtonElement>("reference-add").addEventListener("click", () => open());
  cancel.addEventListener("click", () => dialog.close());
  dialog.addEventListener("keydown", event => { if (event.key === "Escape") event.stopPropagation(); });
  dialog.addEventListener("cancel", event => { if (busy) event.preventDefault(); });
  form.addEventListener("submit", async event => {
    event.preventDefault();
    if (busy) return;
    if (!file.value.trim()) { message.textContent = "Enter the file path first."; file.focus(); return; }
    setBusy(true);
    try {
      await client().saveReference({ id: editing, ...(file.value.trim() !== originalPath ? { path: file.value } : {}), name: name.value, description: description.value });
      dialog.close();
      status.textContent = "Reference saved. It is available to all skills.";
      await refresh();
    } catch (error) { message.textContent = errorText(error); }
    finally { setBusy(false); }
  });
  return { refresh };
}
