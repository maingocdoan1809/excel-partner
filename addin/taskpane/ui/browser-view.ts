/// <reference lib="dom" />
import type { BrowserFrameMessage } from "../../../protocol/src/index.js";

/** Pictures kept for stepping back through a turn. In memory only: never written to history or sent anywhere. */
export const MAX_BROWSER_FRAMES = 20;

export interface BrowserViewElements {
  dock: HTMLElement;
  dockLabel: HTMLElement;
  dockStep: HTMLElement;
  dockUrl: HTMLElement;
  dockImage: HTMLImageElement;
  dockOpen: HTMLButtonElement;
  dockExpand: HTMLButtonElement;
  dockClose: HTMLButtonElement;
  panelLabel: HTMLElement;
  panelStep: HTMLElement;
  panelUrl: HTMLElement;
  panelImage: HTMLImageElement;
  panelPrev: HTMLButtonElement;
  panelNext: HTMLButtonElement;
}

const dataUrl = (frame: BrowserFrameMessage) => `data:${frame.mimeType};base64,${frame.data}`;

/**
 * What the assistant sees in its browser: a small view above the message box that always shows the latest
 * picture, and a larger view where the person can step back through the turn.
 */
export class BrowserView {
  private frames: BrowserFrameMessage[] = [];
  private index = -1;
  /** True while the large view shows the newest picture; stepping back pauses following. */
  private following = true;

  constructor(private readonly els: BrowserViewElements, private readonly openLarge: () => void, private readonly onCount?: (count: number) => void) {
    els.dockOpen.addEventListener("click", openLarge);
    els.dockExpand.addEventListener("click", openLarge);
    els.dockClose.addEventListener("click", () => { els.dock.hidden = true; });
    els.panelPrev.addEventListener("click", () => this.step(-1));
    els.panelNext.addEventListener("click", () => this.step(1));
    els.panelImage.addEventListener("click", () => els.panelImage.parentElement?.classList.toggle("actual"));
    this.reset();
  }

  get count(): number { return this.frames.length; }

  push(frame: BrowserFrameMessage): void {
    this.frames.push(frame);
    if (this.frames.length > MAX_BROWSER_FRAMES) {
      this.frames.shift();
      if (!this.following) this.index = Math.max(0, this.index - 1);
    }
    if (this.following) this.index = this.frames.length - 1;
    this.render();
    this.els.dock.hidden = false;
  }

  /** New turn or new chat: drop every picture. */
  reset(): void {
    this.frames = [];
    this.index = -1;
    this.following = true;
    this.els.dock.hidden = true;
    this.els.dockImage.removeAttribute("src");
    this.els.panelImage.removeAttribute("src");
    this.render();
  }

  private step(delta: number): void {
    if (!this.frames.length) return;
    this.index = Math.min(this.frames.length - 1, Math.max(0, this.index + delta));
    this.following = this.index === this.frames.length - 1;
    this.render();
  }

  private render(): void {
    const { els } = this;
    const latest = this.frames[this.frames.length - 1];
    const shown = this.frames[this.index];
    if (latest) {
      els.dockImage.src = dataUrl(latest);
      els.dockImage.alt = `Browser after: ${latest.label}`;
      els.dockLabel.textContent = latest.label;
      els.dockUrl.textContent = latest.title ? `${latest.title} · ${latest.url ?? ""}`.replace(/ · $/, "") : latest.url ?? "";
      els.dockStep.textContent = `Step ${this.frames.length}`;
    } else {
      els.dockLabel.textContent = "Browser";
      els.dockUrl.textContent = "";
      els.dockStep.textContent = "";
    }
    if (shown) {
      els.panelImage.src = dataUrl(shown);
      els.panelImage.alt = `Browser after: ${shown.label}`;
      els.panelLabel.textContent = shown.label;
      els.panelUrl.textContent = shown.title ? `${shown.title} · ${shown.url ?? ""}`.replace(/ · $/, "") : shown.url ?? "";
      els.panelStep.textContent = `${this.index + 1} / ${this.frames.length}`;
    } else {
      els.panelLabel.textContent = "Nothing yet. Pictures appear when the assistant uses the browser.";
      els.panelUrl.textContent = "";
      els.panelStep.textContent = "";
    }
    els.panelPrev.disabled = this.index <= 0;
    els.panelNext.disabled = this.index < 0 || this.index >= this.frames.length - 1;
    this.onCount?.(this.frames.length);
  }
}
