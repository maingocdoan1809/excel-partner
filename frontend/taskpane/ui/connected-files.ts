import { UNNAMED_FILE, type ConnectedFile } from "../../../protocol/src/index.js";

/** A small picture of the kind of file: a grid for a workbook, text lines for a document, a slide for a presentation. */
const FILE_ICONS: Record<ConnectedFile["host"], string> = {
  excel: '<rect x="3.5" y="4.5" width="13" height="11" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M3.5 8.5h13M3.5 12h13M8 4.5v11" fill="none" stroke="currentColor" stroke-width="1.2"/>',
  word: '<rect x="4.5" y="3.5" width="11" height="13" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M7 7.5h6M7 10h6M7 12.5h4" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/>',
  powerpoint: '<rect x="3" y="4.5" width="14" height="9" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M10 13.5v2.5M7.5 16.5h5" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/>',
};
const APP_NAMES: Record<ConnectedFile["host"], string> = { excel: "Excel", word: "Word", powerpoint: "PowerPoint" };

/**
 * The "Connected files" button in the header and the list it opens: every file open in Word, Excel or PowerPoint with
 * this add-in, the current one first. The green dot shows whether the bridge is connected. The list itself comes from the bridge.
 */
export function setupConnectedFiles(options: {
  /** The list was just opened: ask the bridge for current names (they arrive through `render`). */
  onOpen?: () => void;
} = {}): { render(files: ConnectedFile[]): void; setConnected(connected: boolean): void; close(): void } {
  const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const button = get<HTMLButtonElement>("files-button");
  const popover = get("files-popover");
  const list = get("files-list");
  const hint = get("files-hint");

  const close = () => {
    popover.hidden = true;
    button.setAttribute("aria-expanded", "false");
  };
  const open = () => {
    popover.hidden = false;
    const anchor = button.getBoundingClientRect();
    popover.style.top = `${anchor.bottom + 8}px`;
    popover.style.left = `${Math.max(8, Math.min(anchor.left, window.innerWidth - Math.min(340, window.innerWidth - 16) - 8))}px`;
    button.setAttribute("aria-expanded", "true");
    options.onOpen?.();
  };

  // The button must not stop the click: the Settings menu closes itself on any click outside it, and this one must still
  // see clicks on buttons (like Settings) that stop propagation, so it listens while the event travels down.
  button.addEventListener("click", () => {
    popover.hidden ? open() : close();
  });
  document.addEventListener("click", (event) => {
    const target = event.target as Node;
    if (!popover.hidden && !popover.contains(target) && !button.contains(target)) close();
  }, true);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !popover.hidden) { close(); button.focus(); }
  });

  return {
    close,
    setConnected(connected) {
      button.dataset.connected = String(connected);
    },
    render(files) {
      list.replaceChildren(...files.map((file) => {
        const name = file.name.trim() || UNNAMED_FILE[file.host];
        const row = document.createElement("li");
        row.className = "files-row";
        row.title = name;
        row.dataset.host = file.host;
        row.dataset.current = String(file.current);
        const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        icon.setAttribute("viewBox", "0 0 20 20");
        icon.setAttribute("aria-hidden", "true");
        icon.classList.add("files-icon");
        icon.innerHTML = FILE_ICONS[file.host];
        const label = document.createElement("span");
        label.className = "files-name";
        label.textContent = name;
        const app = document.createElement("span");
        app.className = "files-app";
        app.textContent = APP_NAMES[file.host];
        row.append(icon, label, app);
        if (file.current) {
          const tag = document.createElement("span");
          tag.className = "files-tag";
          tag.textContent = "This file";
          row.append(tag);
        }
        return row;
      }));
      hint.textContent = files.length > 1
        ? "Read or make changes across files. In Ask mode, approve each request."
        : "Open another file with Office Partner to connect it here.";
      button.title = files.length > 1 ? `Connected files (${files.length})` : "Connected files";
    },
  };
}
