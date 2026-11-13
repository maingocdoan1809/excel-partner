/**
 * Keeps a scrolling list pinned to its newest content while the person is at the end of it, and leaves it alone once
 * they scroll up to read. Scrolling back to the end, or `scroll(true)` (for example after they send a message or press
 * a "latest" button), makes it follow again. `onChange` reports each switch so a button can appear and disappear.
 */
export class FollowEnd {
  private following = true;
  private readonly slack: number;
  private readonly onChange: ((following: boolean) => void) | undefined;

  constructor(private readonly element: HTMLElement, options: { slack?: number; onChange?: (following: boolean) => void } = {}) {
    this.slack = options.slack ?? 40;
    this.onChange = options.onChange;
    // Content that grows at the end fires no scroll event, so "following" only changes when the person scrolls.
    element.addEventListener("scroll", () => this.set(this.atEnd()), { passive: true });
  }

  private atEnd(): boolean {
    return this.element.scrollHeight - this.element.scrollTop - this.element.clientHeight <= this.slack;
  }

  private set(following: boolean): void {
    if (following === this.following) return;
    this.following = following;
    this.onChange?.(following);
  }

  /** Move to the end when following. `force` is for what the person asked for themselves. */
  scroll(force = false): void {
    if (force) this.set(true);
    if (this.following) this.element.scrollTop = this.element.scrollHeight;
  }
}

/**
 * Turns "following" changes into a visibility flag for a button without flicker: hidden at once when the person is at
 * the end, shown only if they are still away from it after `delayMs`.
 */
export function steadyVisibility(setVisible: (visible: boolean) => void, delayMs = 250): (following: boolean) => void {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return (following) => {
    clearTimeout(timer);
    if (following) setVisible(false);
    else timer = setTimeout(() => setVisible(true), delayMs);
  };
}
