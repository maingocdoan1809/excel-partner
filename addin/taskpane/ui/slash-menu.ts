import { BUILTIN_COMMANDS, filterCommands, slashQuery, type SlashCommand } from "./slash.js";
import { parseNativeCommand, type ProviderCommand } from "../../../protocol/src/commands.js";

export interface SlashMenuOptions {
  input: HTMLTextAreaElement;
  menu: HTMLElement;
  /** A built-in command was chosen (currently only `clear`). */
  onBuiltin(name: string): void;
  onSelect?(name: string): void;
  getProvider?(): string;
  loadCommands?(): Promise<ProviderCommand[]>;
  onUnavailable?(reason: string): void;
}

/**
 * The `/` menu above the message box: skills from the bridge plus built-in commands. Arrow keys move,
 * Enter/Tab choose, Esc closes; a mouse click also chooses.
 */
export function createSlashMenu({ input, menu, onBuiltin, onSelect, getProvider = () => "", loadCommands, onUnavailable }: SlashMenuOptions) {
  let open = false;
  let items: SlashCommand[] = [];
  let index = 0;
  let skills: SlashCommand[] = [];
  let loadedAt = 0;
  let query = "";
  let loading = false;
  let loadError = false;
  let providerCommands: SlashCommand[] = [];
  let owner = getProvider();
  let generation = 0;
  let pending: Promise<void> | undefined;

  function allCommands(): SlashCommand[] {
    const names = new Set(providerCommands.map(item => item.name));
    return [...providerCommands, ...skills.map(item => names.has(item.name) ? { ...item, name: `skill:${item.name}` } : item),
      ...BUILTIN_COMMANDS.filter(item => !names.has(item.name))];
  }

  async function loadSkills(): Promise<void> {
    // Reload at most every few seconds so edits to a SKILL.md show up without hammering the bridge.
    if (owner !== getProvider()) { owner = getProvider(); generation++; loadedAt = 0; providerCommands = []; pending = undefined; loading = false; }
    if (pending) return pending;
    if (Date.now() - loadedAt < 3_000) return;
    const revision = generation;
    loading = true;
    pending = (async () => { try {
      const native = loadCommands?.().then(commands => ({ commands, error: false }), () => ({ commands: [] as ProviderCommand[], error: true }));
      const response = await fetch("/skills", { cache: "no-store" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const body = (await response.json()) as { skills: Array<{ name: string; description: string }> };
      const catalog = await native;
      if (revision !== generation) return;
      skills = body.skills.map((s) => ({ name: s.name, description: s.description, kind: "skill" as const }));
      providerCommands = catalog?.commands.map(command => ({ name: command.name, description: command.description, kind: "provider", command })) ?? [];
      loadedAt = Date.now();
      loadError = catalog?.error ?? false;
    } catch {
      if (revision !== generation) return;
      skills = [];
      loadError = true;
    } finally {
      if (revision !== generation) return;
      loading = false;
      pending = undefined;
      if (open) render();
    } })();
    return pending;
  }

  function render(): void {
    items = filterCommands(allCommands(), query);
    index = Math.min(index, Math.max(0, items.length - 1));
    menu.hidden = false;
    menu.replaceChildren();

    if (loading || loadError) {
      const status = document.createElement("div");
      status.className = "slash-empty";
      status.textContent = loading ? "Loading commands and skills…" : "Could not load all commands. Check the bridge connection, then type / again.";
      menu.append(status);
    }

    if (items.length === 0 && !loading && !loadError) {
      const empty = document.createElement("div");
      empty.className = "slash-empty";
      empty.textContent = `No command or skill matches /${query}`;
      menu.append(empty);
      return;
    }
    items.forEach((command, i) => {
      const row = document.createElement("div");
      row.className = "slash-item";
      row.setAttribute("role", "option");
      row.setAttribute("aria-selected", String(i === index));
      if (command.command?.execution === "unsupported") row.setAttribute("aria-disabled", "true");
      const head = document.createElement("div");
      const name = document.createElement("span");
      name.className = "slash-name";
      name.textContent = `/${command.name}`;
      const kind = document.createElement("span");
      kind.className = "slash-kind";
      kind.textContent = command.kind === "provider" ? command.command?.execution === "unsupported" ? "CLI only"
        : command.command?.execution === "client" ? "add-in" : getProvider() : command.kind === "skill" ? "bridge skill" : "add-in";
      head.append(name, kind);
      const description = document.createElement("div");
      description.className = "slash-desc";
      description.textContent = [command.description, command.command?.inputHint, command.command?.reason].filter(Boolean).join(" · ");
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
    if (command.command?.execution === "unsupported") {
      onUnavailable?.(command.command.reason ?? "This command requires the CLI terminal.");
      return;
    }
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

  return { update, handleKey, close, reload: () => { generation++; pending = undefined; loading = false; loadedAt = 0; providerCommands = []; if (open) void loadSkills(); },
    async resolve(text: string) {
      const parsed = parseNativeCommand(text);
      if (!parsed) return undefined;
      await loadSkills();
      if (loadError && !allCommands().some(item => item.name === parsed.name)) throw new Error("Could not verify this command. Reopen the / menu and try again.");
      return allCommands().find(item => item.name === parsed.name);
    },
    get commands() { return allCommands(); }, get isOpen() { return open; } };
}
