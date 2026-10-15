import { BUILTIN_COMMANDS, filterCommands, slashQuery, type SlashCommand } from "./slash.js";

export interface SlashMenuOptions {
  input: HTMLTextAreaElement;
  menu: HTMLElement;
  /** A built-in command was chosen (currently only `clear`). */
  onBuiltin(name: string): void;
  onSelect?(name: string): void;
}

/**
 * The `/` menu above the message box: skills from the bridge plus built-in commands. Arrow keys move,
 * Enter/Tab choose, Esc closes; a mouse click also chooses.
 */
export function createSlashMenu({ input, menu, onBuiltin, onSelect }: SlashMenuOptions) {
  let open = false;
  let items: SlashCommand[] = [];
  let index = 0;
  let skills: SlashCommand[] = [];
  let loadedAt = 0;
  let query = "";
  let loading = false;
  let loadError = false;

  async function loadSkills(): Promise<void> {
    // Reload at most every few seconds so edits to a SKILL.md show up without hammering the bridge.
    if (loading || Date.now() - loadedAt < 3_000) return;
    loading = true;
    try {
      const response = await fetch("/skills", { cache: "no-store" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const body = (await response.json()) as { skills: Array<{ name: string; description: string }> };
      skills = body.skills.map((s) => ({ name: s.name, description: s.description, kind: "skill" as const }));
      loadedAt = Date.now();
      loadError = false;
    } catch {
      skills = [];
      loadError = true;
    } finally {
      loading = false;
      if (open) render();
    }
  }

  function render(): void {
    items = filterCommands([...skills, ...BUILTIN_COMMANDS], query);
    index = Math.min(index, Math.max(0, items.length - 1));
    menu.hidden = false;
    menu.replaceChildren();

    if (loading || loadError) {
      const status = document.createElement("div");
      status.className = "slash-empty";
      status.textContent = loading ? "Loading skills…" : "Could not load skills. Check the bridge connection, then type / again.";
      menu.append(status);
    }

    if (items.length === 0 && !loading && !loadError) {
      const empty = document.createElement("div");
      empty.className = "slash-empty";
      empty.textContent = skills.length === 0 ? "No skills found. Add a folder with a SKILL.md under skills/." : `No skill matches /${query}`;
      menu.append(empty);
      return;
    }
    items.forEach((command, i) => {
      const row = document.createElement("div");
      row.className = "slash-item";
      row.setAttribute("role", "option");
      row.setAttribute("aria-selected", String(i === index));
      const head = document.createElement("div");
      const name = document.createElement("span");
      name.className = "slash-name";
      name.textContent = `/${command.name}`;
      const kind = document.createElement("span");
      kind.className = "slash-kind";
      kind.textContent = command.kind === "skill" ? "skill" : "command";
      head.append(name, kind);
      const description = document.createElement("div");
      description.className = "slash-desc";
      description.textContent = command.description;
      row.append(head, description);
      // mousedown (not click): the textarea's blur would close the menu first.
      row.addEventListener("mousedown", (event) => {
        event.preventDefault();
        choose(command);
      });
      menu.append(row);
      if (i === index) row.scrollIntoView({ block: "nearest" });
    });
  }

  function close(): void {
    open = false;
    index = 0;
    menu.hidden = true;
  }

  function choose(command: SlashCommand): void {
    if (command.kind === "builtin") {
      input.value = "";
      close();
      onBuiltin(command.name);
    } else {
      input.value = `/${command.name} `;
      onSelect?.(command.name);
      close();
    }
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  /** Call on every `input` event. */
  function update(): void {
    const q = slashQuery(input.value);
    if (q === null) return close();
    query = q;
    open = true;
    void loadSkills().then(() => {
      // The user may have kept typing while the list loaded.
      if (open && slashQuery(input.value) === query) render();
    });
    render();
  }

  /** Returns true when the key was consumed by the menu. */
  function handleKey(event: KeyboardEvent): boolean {
    if (!open) return false;
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      if (items.length === 0) return true;
      index = (index + (event.key === "ArrowDown" ? 1 : -1) + items.length) % items.length;
      event.preventDefault();
      render();
      return true;
    }
    if ((event.key === "Enter" && !event.ctrlKey && !event.metaKey) || event.key === "Tab") {
      const command = items[index];
      if (!command) return false;
      event.preventDefault();
      choose(command);
      return true;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      close();
      return true;
    }
    return false;
  }

  return { update, handleKey, close, reload: () => { loadedAt = 0; void loadSkills(); }, get isOpen() { return open; } };
}
