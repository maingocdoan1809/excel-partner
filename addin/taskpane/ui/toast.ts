let current: { element: HTMLElement; timer: number } | null = null;

/** A short message that floats over the chat and disappears by itself. Not saved to the chat history. */
export function showToast(text: string, durationMs = Math.max(6000, text.length * 70)): void {
  if (current) { window.clearTimeout(current.timer); current.element.remove(); }
  const element = document.createElement("div");
  element.className = "toast";
  element.setAttribute("role", "alert");
  element.textContent = text;
  element.addEventListener("click", () => { window.clearTimeout(timer); element.remove(); if (current?.element === element) current = null; });
  document.body.append(element);
  const timer = window.setTimeout(() => { element.remove(); if (current?.element === element) current = null; }, durationMs);
  current = { element, timer };
}
