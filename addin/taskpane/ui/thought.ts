import { toolChevron } from "./tool-view.js";

interface Disclosure {
  element: HTMLDetailsElement;
  label: HTMLElement;
  body: HTMLElement;
  text: string;
  /** When the current run of reasoning began; null while the answer or a tool call is what is happening. */
  startedAt: number | null;
  elapsedMs: number;
}

/**
 * The model's reasoning for the running turn, in one disclosure above its answer: open while the model thinks, folded
 * (with how long it took) once an answer or a tool call follows. Shown live only; it is not saved with the chat.
 */
export class ThoughtView {
  private current: Disclosure | null = null;

  /** `before`: the element the disclosure goes above (the working indicator), if there is one. */
  constructor(private readonly container: HTMLElement, private readonly before: () => HTMLElement | null, private readonly scroll: () => void) {}

  /** Add reasoning text. The first text opens the disclosure; text after a pause starts a new paragraph. */
  add(text: string): void {
    if (!text) return;
    const thought = this.current ?? this.create();
    if (thought.startedAt === null) {
      if (thought.text) thought.text += "\n\n";
      thought.startedAt = Date.now();
      thought.label.textContent = "Thinking…";
    }
    thought.text = (thought.text + text).replace(/\n{3,}/g, "\n\n");
    thought.body.textContent = thought.text.trimStart();
    thought.body.scrollTop = thought.body.scrollHeight;
    this.scroll();
  }

  /** The answer or a tool call follows the reasoning: fold the disclosure and say how long it took. */
  end(): void {
    const thought = this.current;
    if (!thought || thought.startedAt === null) return;
    thought.elapsedMs += Date.now() - thought.startedAt;
    thought.startedAt = null;
    thought.label.textContent = `Thought for ${Math.max(1, Math.round(thought.elapsedMs / 1000))}s`;
    thought.element.open = false;
  }

  /** The turn is over (or the chat changed): the next reasoning gets a disclosure of its own. */
  reset(): void {
    this.end();
    this.current = null;
  }

  private create(): Disclosure {
    // The same look as the "N tools called" disclosure.
    const element = document.createElement("details");
    element.className = "tool-group thought";
    element.open = true;
    const summary = document.createElement("summary");
    const label = document.createElement("span");
    const chevron = document.createElement("span");
    chevron.className = "tool-group-chevron";
    chevron.append(toolChevron());
    summary.append(label, chevron);
    const body = document.createElement("div");
    body.className = "thought-body";
    element.append(summary, body);
    this.container.querySelector(".welcome")?.remove();
    const before = this.before();
    if (before && before.parentElement === this.container) this.container.insertBefore(element, before);
    else this.container.append(element);
    this.current = { element, label, body, text: "", startedAt: null, elapsedMs: 0 };
    return this.current;
  }
}
