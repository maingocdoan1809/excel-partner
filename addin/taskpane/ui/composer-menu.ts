export function setupComposerMenu(select: HTMLSelectElement, button: HTMLButtonElement, menu: HTMLElement, label: (value: string) => string = (value) => value): { refresh(): void; close(): void } {
  const close = () => {
    menu.hidden = true;
    button.setAttribute("aria-expanded", "false");
  };
  const refresh = () => {
    const selected = select.selectedOptions[0];
    button.querySelector(".picker-label")!.textContent = label(selected?.textContent ?? "Default");
    button.title = select.title;
    menu.replaceChildren(...Array.from(select.options).map((option) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "picker-option";
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", String(option.selected));
      item.disabled = option.disabled;
      item.textContent = option.textContent;
      item.addEventListener("click", () => {
        select.value = option.value;
        select.dispatchEvent(new Event("change", { bubbles: true }));
        refresh();
        close();
        button.focus();
      });
      return item;
    }));
  };
  const open = () => {
    document.dispatchEvent(new CustomEvent("composer-picker-open", { detail: button.id }));
    refresh();
    menu.hidden = false;
    button.setAttribute("aria-expanded", "true");
    const selected = menu.querySelector<HTMLButtonElement>('[aria-selected="true"]:not(:disabled)');
    (selected ?? menu.querySelector<HTMLButtonElement>("button:not(:disabled)"))?.focus();
  };
  button.addEventListener("click", () => menu.hidden ? open() : close());
  button.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown" || event.key === "ArrowUp") { event.preventDefault(); open(); }
  });
  menu.addEventListener("keydown", (event) => {
    if (event.key === "Escape" || event.key === "Tab") {
      close();
      if (event.key === "Escape") { event.preventDefault(); button.focus(); }
      return;
    }
    if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return;
    event.preventDefault();
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button:not(:disabled)"));
    const current = items.indexOf(document.activeElement as HTMLButtonElement);
    const next = event.key === "Home" ? 0 : event.key === "End" ? items.length - 1 : (current + (event.key === "ArrowDown" ? 1 : -1) + items.length) % items.length;
    items[next]?.focus();
  });
  document.addEventListener("pointerdown", event => {
    if (!button.contains(event.target as Node) && !menu.contains(event.target as Node)) close();
  });
  document.addEventListener("composer-picker-open", event => {
    if ((event as CustomEvent<string>).detail !== button.id) close();
  });
  select.addEventListener("change", refresh);
  const observer = new MutationObserver(refresh);
  observer.observe(select, { childList: true, subtree: true, characterData: true });
  window.addEventListener("pagehide", () => observer.disconnect(), { once: true });
  refresh();
  return { refresh, close };
}
