let current: { element: HTMLElement; timer: number } | null = null;

const ICON = '<svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="8" fill="none" stroke="currentColor" stroke-width="1.5"/><path d="M10 5.8v5M10 13.4v.1" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>';

/** Slide the toast out, then remove it. */
function dismiss(toast: { element: HTMLElement; timer: number }): void {
  window.clearTimeout(toast.timer);
  if (current === toast) current = null;
  toast.element.classList.add("leaving");
  const remove = () => toast.element.remove();
  toast.element.addEventListener("animationend", remove, { once: true });
  window.setTimeout(remove, 400); // when animations are turned off
}

/** A short message that slides in at the top of the pane and goes away by itself. Not saved to the chat history. */
export function showToast(text: string, durationMs = Math.max(6000, text.length * 70)): void {
  if (current) { window.clearTimeout(current.timer); current.element.remove(); current = null; }
  const element = document.createElement("div");
  element.className = "toast";
  element.setAttribute("role", "alert");
  element.style.setProperty("--toast-time", `${durationMs}ms`);
  const icon = document.createElement("span"); icon.className = "toast-icon"; icon.innerHTML = ICON;
  const message = document.createElement("span"); message.className = "toast-text"; message.textContent = text;
  const close = document.createElement("button"); close.type = "button"; close.className = "toast-close"; close.setAttribute("aria-label", "Dismiss"); close.textContent = "×";
  const bar = document.createElement("span"); bar.className = "toast-bar";
  element.append(icon, message, close, bar);
  const toast = { element, timer: 0 };
  close.addEventListener("click", () => dismiss(toast));
  toast.timer = window.setTimeout(() => dismiss(toast), durationMs);
  current = toast;
  document.body.append(element);
}
