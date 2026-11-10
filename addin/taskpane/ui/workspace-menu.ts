import type { BridgeClient, SavedReference, UploadedAttachment, WorkspaceEntry, WorkspaceRoot } from "../bridge-client.js";
import { activeHost } from "../host/index.js";
import { activeFileMention } from "./file-mentions.js";
import type { StoredMessage } from "./sessions.js";
import { filesFromDrop, isFileDrag } from "./transfer-files.js";

export interface WorkspaceMenuDeps {
  getClient(): BridgeClient;
  getWorkbookName(): Promise<string>;
  closeSlash(): void;
  renderComposerPreview(): void;
  append(message: StoredMessage): HTMLElement;
  attachments: Array<{ reference: UploadedAttachment; name: string }>;
  elements: { input: HTMLTextAreaElement; mentionMenu: HTMLElement };
}

/** The @ menu, granted workspace folders, and local files chosen or dropped into the composer. */
export function setupWorkspaceMenu(deps: WorkspaceMenuDeps) {
  const els = deps.elements;
  const pendingAttachments = deps.attachments;
  let workspaceRoots: WorkspaceRoot[] | null = null;
  let mentionReferences: SavedReference[] = [];
  let workspaceStack: { path: string; name: string }[] = [];
  let workspaceEntries: WorkspaceEntry[] = [];
  let workspaceError: string | null = null;
  let workspaceDefaultError: string | null = null;
  let workspaceLoading = false;
  let workspaceRequest = 0;
  let localPickerLoading = false;

  function updateMentionMenu(): void {
    const before = els.input.value.slice(0, els.input.selectionStart);
    const mention = activeFileMention(before, pendingAttachments.map((item) => item.name));
    const show = !!mention;
    const opening = show && els.mentionMenu.hidden;
    els.mentionMenu.hidden = !show;
    if (show) {
      deps.closeSlash();
      void ensureWorkspaces();
      if (opening) void refreshMentionReferences();
      renderWorkspaceMenu();
    }
  }

  /** Saved reference files are listed in the @ menu; reload them each time it opens so edits show up. */
  async function refreshMentionReferences(): Promise<void> {
    try { mentionReferences = await deps.getClient().listSavedReferences(); }
    catch { mentionReferences = []; }
    renderWorkspaceMenu();
  }

  async function ensureWorkspaces(): Promise<void> {
    if (workspaceRoots || workspaceLoading) return;
    workspaceLoading = true;
    workspaceError = null;
    renderWorkspaceMenu();
    try {
      const [documentUrl, workbookName] = await Promise.all([currentWorkbookUrl(), activeHost().hasSheets ? deps.getWorkbookName() : Promise.resolve("")]);
      const result = await deps.getClient().listWorkspaces(documentUrl, workbookName);
      workspaceRoots = result.roots;
      workspaceDefaultError = result.defaultError;
    } catch (error) {
      workspaceError = error instanceof Error ? error.message : String(error);
    } finally {
      workspaceLoading = false;
      renderWorkspaceMenu();
    }
  }

  async function currentWorkbookUrl(): Promise<string> {
    const direct = Office.context.document.url;
    if (direct) return direct;
    return new Promise((resolve) => {
      const timer = window.setTimeout(() => resolve(""), 2500);
      try {
        Office.context.document.getFilePropertiesAsync((result) => {
          window.clearTimeout(timer);
          resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value.url ?? "" : "");
        });
      } catch { window.clearTimeout(timer); resolve(""); }
    });
  }

  async function openWorkspaceFolder(path: string, name: string): Promise<void> {
    const before = els.input.value.slice(0, els.input.selectionStart);
    const mention = activeFileMention(before, pendingAttachments.map((item) => item.name));
    if (mention) {
      const cursor = mention.start + 1;
      els.input.value = `${before.slice(0, cursor)}${els.input.value.slice(els.input.selectionStart)}`;
      els.input.setSelectionRange(cursor, cursor);
      deps.renderComposerPreview();
    }
    workspaceStack.push({ path, name });
    const request = ++workspaceRequest;
    workspaceLoading = true;
    workspaceError = null;
    renderWorkspaceMenu();
    try {
      const listing = await deps.getClient().listWorkspaceEntries(path);
      if (request === workspaceRequest) workspaceEntries = listing.entries;
    } catch (error) {
      if (request === workspaceRequest) workspaceError = error instanceof Error ? error.message : String(error);
    } finally {
      if (request === workspaceRequest) {
        workspaceLoading = false;
        renderWorkspaceMenu();
      }
    }
  }

  function renderWorkspaceMenu(): void {
    if (els.mentionMenu.hidden) return;
    const menu = els.mentionMenu;
    menu.replaceChildren();
    const heading = document.createElement("div");
    heading.className = "workspace-heading";
    if (workspaceStack.length) {
      const back = document.createElement("button");
      back.type = "button";
      back.textContent = "←";
      back.title = "Back";
      back.addEventListener("click", () => {
        workspaceStack.pop();
        if (workspaceStack.length) {
          const parent = workspaceStack.pop()!;
          void openWorkspaceFolder(parent.path, parent.name);
        }
        else renderWorkspaceMenu();
      });
      heading.append(back);
    }
    const title = document.createElement("strong");
    title.textContent = workspaceStack.length ? workspaceStack.map((part) => part.name).join(" / ") : "Workspaces";
    title.title = workspaceStack[workspaceStack.length - 1]?.path ?? "Workspaces";
    heading.append(title);
    menu.append(heading);

    if (workspaceLoading) menu.append(workspaceNote("Loading…"));
    else if (workspaceError) menu.append(workspaceNote(workspaceError));
    else {
      const rows = workspaceStack.length ? workspaceEntries : workspaceRoots ?? [];
      const before = els.input.value.slice(0, els.input.selectionStart);
      const query = activeFileMention(before, pendingAttachments.map((item) => item.name))?.query.toLowerCase().trim() ?? "";
      const visible = workspaceStack.length && query ? rows.filter((entry) => entry.name.toLowerCase().includes(query)) : rows;
      if (!visible.length) menu.append(workspaceNote(workspaceStack.length ? "No matching files or folders." : "No workbook folder found. Choose local files below."));
      for (const entry of visible) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "workspace-row";
        button.title = entry.path;
        const isFolder = !workspaceStack.length || ("type" in entry && entry.type === "folder");
        const supported = isFolder || ("supported" in entry && entry.supported);
        setWorkspaceRow(button, `${entry.name}${!workspaceStack.length && "isDefault" in entry && entry.isDefault ? " · current workbook" : ""}${!supported ? " · unsupported" : ""}`, isFolder);
        button.disabled = !supported;
        button.addEventListener("click", () => {
          if (isFolder) void openWorkspaceFolder(entry.path, entry.name);
          else void chooseWorkspaceFile(entry.path);
        });
        if (!workspaceStack.length && "removable" in entry && entry.removable) {
          const line = document.createElement("div");
          line.className = "workspace-root";
          const remove = document.createElement("button");
          remove.type = "button";
          remove.className = "workspace-remove";
          remove.textContent = "×";
          remove.title = `Remove workspace ${entry.name}`;
          remove.setAttribute("aria-label", `Remove workspace ${entry.name}`);
          remove.addEventListener("click", () => void removeWorkspace(entry.path));
          line.append(button, remove);
          menu.append(line);
        } else menu.append(button);
      }
      if (!workspaceStack.length && workspaceDefaultError && visible.length) menu.append(workspaceNote("Workbook folder unavailable. Choose local files below."));
    }

    // The saved references (files and folders) are always offered, even when the workspace folders could not be listed.
    if (!workspaceStack.length && !workspaceLoading) {
      const before = els.input.value.slice(0, els.input.selectionStart);
      const query = activeFileMention(before, pendingAttachments.map((item) => item.name))?.query.toLowerCase().trim() ?? "";
      const refs = query ? mentionReferences.filter((ref) => `${ref.name} ${ref.path}`.toLowerCase().includes(query)) : mentionReferences;
      if (refs.length) {
        const label = document.createElement("div");
        label.className = "workspace-heading";
        label.textContent = "References";
        menu.append(label);
        for (const ref of refs) {
          const kind = ref.kind === "folder" ? "Folder" : "File";
          const button = document.createElement("button");
          button.type = "button";
          button.className = "workspace-row";
          button.title = ref.available ? `${ref.path}${ref.description ? `\n${ref.description}` : ""}` : `${kind} not found: ${ref.path}`;
          setWorkspaceRow(button, `${ref.name}${ref.available ? "" : ` · ${kind.toLowerCase()} not found`}`, ref.kind === "folder");
          button.disabled = !ref.available;
          // A file is attached by its path, like any workspace file. A folder is only pointed at: its path goes to the assistant,
          // which lists and reads inside it itself, so there is nothing to open or attach here.
          button.addEventListener("click", () => ref.kind === "folder"
            ? addAttachment({ reference: { name: ref.name, kind: "folder", path: ref.path }, name: ref.name })
            : void chooseWorkspaceFile(ref.path));
          menu.append(button);
        }
      }
    }

    const actions = document.createElement("div");
    actions.className = "workspace-actions";
    const pick = document.createElement("button");
    pick.type = "button";
    setWorkspaceRow(pick, "Choose local files…", false);
    pick.addEventListener("click", openLocalFilePicker);
    actions.append(pick);
    menu.append(actions);
  }

  function setWorkspaceRow(button: HTMLButtonElement, text: string, folder: boolean): void {
    const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    icon.setAttribute("viewBox", "0 0 20 20");
    icon.setAttribute("aria-hidden", "true");
    icon.classList.add("workspace-icon");
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", folder ? "M2.5 6V4.5h6l2 2h7v10h-15z" : "M5 2.5h6l4 4v11H5z M11 2.5v4h4 M7.5 10h5 M7.5 13h5");
    icon.append(path);
    const label = document.createElement("span");
    label.className = "workspace-row-label";
    label.textContent = text;
    button.replaceChildren(icon, label);
  }

  async function removeWorkspace(path: string): Promise<void> {
    workspaceLoading = true;
    workspaceError = null;
    renderWorkspaceMenu();
    try {
      await deps.getClient().removeWorkspace(path);
      workspaceRoots = workspaceRoots?.filter((root) => root.path !== path) ?? null;
      workspaceStack = [];
    } catch (error) {
      workspaceError = error instanceof Error ? error.message : String(error);
    } finally {
      workspaceLoading = false;
      renderWorkspaceMenu();
    }
  }

  function workspaceNote(message: string): HTMLElement {
    const note = document.createElement("div");
    note.className = "slash-empty";
    note.textContent = message;
    return note;
  }

  async function chooseWorkspaceFile(path: string): Promise<void> {
    workspaceLoading = true;
    renderWorkspaceMenu();
    try {
      const attachment = await deps.getClient().attachWorkspaceFile(path);
      selectWorkspaceAttachment(attachment);
    } catch (error) {
      workspaceError = error instanceof Error ? error.message : String(error);
    } finally {
      workspaceLoading = false;
      renderWorkspaceMenu();
    }
  }

  function selectWorkspaceAttachment(attachment: UploadedAttachment): void {
    addAttachment({ reference: attachment, name: attachment.name });
  }

  function openLocalFilePicker(): void {
    if (localPickerLoading) return;
    const input = document.createElement("input");
    input.type = "file";
    input.multiple = true;
    input.onchange = () => {
      const files = Array.from(input.files ?? []);
      if (files.length) void chooseLocalFiles(files);
    };
    input.click();
  }

  /**
   * Drop files anywhere on the pane to attach them. The whole window listens because a file dropped outside
   * the message box would otherwise make the browser open it and replace the pane.
   */
  function setupFileDrop(): void {
    const card = els.input.closest<HTMLElement>(".composer-card");
    let depth = 0;
    const reset = () => { depth = 0; card?.classList.remove("drop-active"); };
    const composerVisible = () => els.input.getClientRects().length > 0;
    document.addEventListener("dragenter", (event) => {
      if (!isFileDrag(event.dataTransfer)) return;
      event.preventDefault();
      depth++;
      if (composerVisible()) card?.classList.add("drop-active");
    });
    document.addEventListener("dragover", (event) => {
      if (!isFileDrag(event.dataTransfer)) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = composerVisible() ? "copy" : "none";
    });
    document.addEventListener("dragleave", (event) => {
      if (!isFileDrag(event.dataTransfer)) return;
      depth = Math.max(0, depth - 1);
      if (!depth) reset();
    });
    document.addEventListener("dragend", reset);
    document.addEventListener("drop", (event) => {
      if (!isFileDrag(event.dataTransfer)) return;
      event.preventDefault();
      reset();
      if (!composerVisible()) return;
      const { files, folders } = filesFromDrop(event.dataTransfer);
      if (folders) deps.append({ kind: "error", text: "Folders cannot be attached. Drop the files inside the folder instead." });
      if (files.length) void chooseLocalFiles(files);
    });
  }

  async function chooseLocalFiles(files: File[]): Promise<void> {
    if (localPickerLoading) return;
    if (pendingAttachments.length + files.length > 10) {
      deps.append({ kind: "error", text: "Choose at most 10 files per message." });
      return;
    }
    localPickerLoading = true;
    els.mentionMenu.hidden = true;
    try {
      for (const file of files) selectWorkspaceAttachment(await deps.getClient().uploadAttachment(file));
    } catch (error) {
      deps.append({ kind: "error", text: error instanceof Error ? error.message : String(error) });
    } finally {
      localPickerLoading = false;
    }
  }

  function addAttachment(item: { reference: UploadedAttachment; name: string }): void {
    if (pendingAttachments.length >= 10) {
      deps.append({ kind: "error", text: "Choose at most 10 files per message." });
      return;
    }
    els.mentionMenu.hidden = true;
    const cursor = els.input.selectionStart;
    const before = els.input.value.slice(0, cursor);
    const after = els.input.value.slice(cursor);
    const token = `@${item.name}`;
    const mention = activeFileMention(before, pendingAttachments.map((selected) => selected.name));
    const prefix = mention ? before.slice(0, mention.start) : `${before}${before && !/\s$/.test(before) ? " " : ""}`;
    const separator = !after || !/^\s/.test(after) ? " " : "";
    els.input.value = `${prefix}${token}${separator}${after}`;
    const nextCursor = prefix.length + token.length + 1;
    els.input.setSelectionRange(nextCursor, nextCursor);
    pendingAttachments.push(item);
    deps.renderComposerPreview();
    els.input.focus();
  }

  function resetWorkspaces(): void {
    workspaceRoots = null;
    workspaceStack = [];
  }

  return { updateMentionMenu, ensureWorkspaces, resetWorkspaces, chooseLocalFiles, setupFileDrop };
}
