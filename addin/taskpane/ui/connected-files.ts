import { UNNAMED_WORKBOOK, type ConnectedFile } from "../../../protocol/src/index.js";

/**
 * The "Connected files" button in the header and the list it opens: every workbook open in Excel with this add-in,
 * the current one first. The green dot shows whether the bridge is connected. The list itself comes from the bridge.
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
        const name = file.name.trim() || UNNAMED_WORKBOOK;
        const row = document.createElement("li");
        row.className = "files-row";
        row.title = name;
        row.dataset.current = String(file.current);
        const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        icon.setAttribute("viewBox", "0 0 20 20");
        icon.setAttribute("aria-hidden", "true");
        icon.classList.add("files-icon");
        icon.innerHTML = '<rect x="3.5" y="4.5" width="13" height="11" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M3.5 8.5h13M3.5 12h13M8 4.5v11" fill="none" stroke="currentColor" stroke-width="1.2"/>';
        const label = document.createElement("span");
        label.className = "files-name";
        label.textContent = name;
        row.append(icon, label);
        if (file.current) {
          const tag = document.createElement("span");
          tag.className = "files-tag";
          tag.textContent = "This file";
          row.append(tag);
        }
        return row;
      }));
      hint.textContent = files.length > 1
        ? "The assistant can read these files when you ask, and can ask their assistants to make changes in them. In Ask mode you approve each request first."
        : "Open another workbook with Excel Partner and the assistant can read it, or ask its assistant to change it, from here.";
      button.title = files.length > 1 ? `Connected files (${files.length})` : "Connected files";
    },
  };
}
