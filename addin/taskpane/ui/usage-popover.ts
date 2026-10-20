export function setupUsagePopover(button: HTMLButtonElement, popover: HTMLElement, refresh: () => void): void {
  const close = () => {
    popover.hidden = true;
    button.setAttribute("aria-expanded", "false");
  };
  button.addEventListener("click", () => {
    if (!popover.hidden) return close();
    document.dispatchEvent(new CustomEvent("composer-picker-open", { detail: button.id }));
    popover.hidden = false;
    button.setAttribute("aria-expanded", "true");
    refresh();
  });
  document.addEventListener("pointerdown", event => {
    if (!button.contains(event.target as Node) && !popover.contains(event.target as Node)) close();
  });
  document.addEventListener("keydown", event => {
    if (!popover.hidden && event.key === "Escape") {
      event.preventDefault();
      close();
      button.focus();
    }
  });
  document.addEventListener("composer-picker-open", event => {
    if ((event as CustomEvent<string>).detail !== button.id) close();
  });
}
