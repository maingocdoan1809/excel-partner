// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ThoughtView } from "../frontend/taskpane/ui/thought.js";

let messages: HTMLElement;
let dots: HTMLElement | null;
let scrolled: number;
let view: ThoughtView;
beforeEach(() => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date("2026-11-05T12:00:00Z"));
  document.body.innerHTML = '<div id="messages"><div class="welcome"></div></div>';
  messages = document.getElementById("messages")!;
  dots = document.createElement("div");
  dots.className = "thinking";
  messages.append(dots);
  scrolled = 0;
  view = new ThoughtView(messages, () => dots, () => { scrolled++; });
});
afterEach(() => { vi.useRealTimers(); document.body.replaceChildren(); });

const disclosure = () => messages.querySelector<HTMLDetailsElement>("details.thought");
const label = () => disclosure()!.querySelector("summary > span")!.textContent;
const body = () => disclosure()!.querySelector(".thought-body")!.textContent;

describe("reasoning disclosure", () => {
  it("opens above the working indicator on the first text and grows as text streams", () => {
    view.add("Check the ");
    view.add("header row.");
    expect(messages.querySelector(".welcome")).toBeNull();
    expect(disclosure()!.open).toBe(true);
    expect(disclosure()!.nextElementSibling).toBe(dots);
    expect(label()).toBe("Thinking…");
    expect(body()).toBe("Check the header row.");
    expect(messages.querySelectorAll("details.thought")).toHaveLength(1);
    expect(scrolled).toBe(2);
  });

  it("folds and says how long it took once the answer or a tool call follows", () => {
    view.add("Plan.");
    vi.advanceTimersByTime(4_400);
    view.end();
    expect(disclosure()!.open).toBe(false);
    expect(label()).toBe("Thought for 4s");
    view.end();
    expect(label()).toBe("Thought for 4s");
  });

  it("never says less than a second", () => {
    view.add("Quick.");
    view.end();
    expect(label()).toBe("Thought for 1s");
  });

  it("continues in the same disclosure after a tool call, in a new paragraph, adding up the time", () => {
    view.add("First.");
    vi.advanceTimersByTime(2_000);
    view.end();
    vi.advanceTimersByTime(30_000); // a tool call: not thinking time
    view.add("Second.");
    expect(messages.querySelectorAll("details.thought")).toHaveLength(1);
    expect(label()).toBe("Thinking…");
    expect(body()).toBe("First.\n\nSecond.");
    vi.advanceTimersByTime(3_000);
    view.end();
    expect(label()).toBe("Thought for 5s");
  });

  it("does not stack blank lines and does not start with one", () => {
    view.add("\n\nOne.");
    view.add("\n\n\n\nTwo.");
    expect(body()).toBe("One.\n\nTwo.");
  });

  it("gives the next turn its own disclosure after reset, leaving the finished one in place", () => {
    view.add("Turn one.");
    view.reset();
    view.add("Turn two.");
    const all = messages.querySelectorAll<HTMLDetailsElement>("details.thought");
    expect(all).toHaveLength(2);
    expect(all[0]!.querySelector(".thought-body")!.textContent).toBe("Turn one.");
    expect(all[0]!.open).toBe(false);
    expect(all[1]!.querySelector(".thought-body")!.textContent).toBe("Turn two.");
  });

  it("shows nothing for empty text and appends at the end when there is no working indicator", () => {
    view.add("");
    expect(disclosure()).toBeNull();
    dots!.remove();
    dots = null;
    view.add("Late.");
    expect(messages.lastElementChild).toBe(disclosure());
  });

  it("renders text as text, never as markup", () => {
    view.add("<img src=x onerror=alert(1)> & <b>bold</b>");
    expect(disclosure()!.querySelector("img, b")).toBeNull();
    expect(body()).toBe("<img src=x onerror=alert(1)> & <b>bold</b>");
  });
});
