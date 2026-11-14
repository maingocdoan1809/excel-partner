/** What the AI is doing now, shown at the end of the chat: moving dots, words for the current step and how long that step has taken. */
export function setupThinking(messages: HTMLElement, hooks: { scrollToEnd(): void; onClear(): void }) {
  let element: HTMLElement | null = null;
  let timer: number | undefined;
  let stepStartedAt = 0;
  /** Seconds are shown only once a step is slow enough for the wait to be noticed. */
  const SHOW_ELAPSED_AFTER_S = 3;

  const tick = () => {
    const elapsed = element?.querySelector<HTMLElement>(".thinking-elapsed");
    if (!elapsed) return;
    const seconds = Math.floor((Date.now() - stepStartedAt) / 1000);
    elapsed.textContent = seconds >= SHOW_ELAPSED_AFTER_S ? `${seconds}s` : "";
  };
  const stopTimer = () => { if (timer !== undefined) window.clearInterval(timer); timer = undefined; };

  return {
    /** The dots while they are shown, so new messages can be put before them. */
    get element(): HTMLElement | null { return element; },
    show(): void {
      if (element) return;
      element = document.createElement("div");
      element.className = "thinking";
      element.setAttribute("role", "status");
      element.setAttribute("aria-label", "AI is working");
      element.innerHTML = '<span class="thinking-dots" aria-hidden="true"><i></i><i></i><i></i></span>';
      messages.append(element);
      stepStartedAt = Date.now();
      timer = window.setInterval(tick, 1000);
      hooks.scrollToEnd();
    },
    /**
     * Words next to the dots for the step now running ("Thinking…", "Using read range…", a provider's own wait message); null goes back to
     * dots only. A new step starts its own count of seconds.
     */
    setLabel(text: string | null): void {
      if (!element) return;
      const current = element.querySelector(".thinking-label")?.textContent ?? null;
      if (current === text) return;
      element.querySelector(".thinking-label")?.remove();
      element.querySelector(".thinking-elapsed")?.remove();
      element.setAttribute("aria-label", text ?? "AI is working");
      stepStartedAt = Date.now();
      if (!text) return;
      const label = document.createElement("span");
      label.className = "thinking-label";
      label.textContent = text;
      const elapsed = document.createElement("span");
      elapsed.className = "thinking-elapsed";
      elapsed.setAttribute("aria-hidden", "true");
      element.append(label, elapsed);
    },
    clear(): void {
      hooks.onClear();
      stopTimer();
      element?.remove();
      element = null;
    },
  };
}
