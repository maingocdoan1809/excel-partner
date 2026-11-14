// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { FollowEnd, steadyVisibility } from "../frontend/taskpane/ui/follow-end.js";

/** A list with a given content height; scrolling is simulated the way a browser reports it. */
function list(clientHeight = 300) {
  const element = document.createElement("div");
  let scrollHeight = 1000;
  Object.defineProperty(element, "clientHeight", { value: clientHeight });
  Object.defineProperty(element, "scrollHeight", { get: () => scrollHeight });
  const userScrollsTo = (top: number) => { element.scrollTop = top; element.dispatchEvent(new Event("scroll")); };
  return { element, grow: (by: number) => { scrollHeight += by; }, userScrollsTo };
}

describe("FollowEnd", () => {
  it("follows new content while the person is at the end", () => {
    const { element, grow } = list();
    const follow = new FollowEnd(element);
    follow.scroll();
    expect(element.scrollTop).toBe(1000);
    grow(200);
    follow.scroll();
    expect(element.scrollTop).toBe(1200);
  });

  it("leaves the person where they are once they scroll up", () => {
    const { element, grow, userScrollsTo } = list();
    const follow = new FollowEnd(element);
    follow.scroll();
    userScrollsTo(100);
    grow(500);
    follow.scroll();
    expect(element.scrollTop).toBe(100);
    follow.scroll();
    expect(element.scrollTop).toBe(100);
  });

  it("follows again when the person scrolls back to the end", () => {
    const { element, grow, userScrollsTo } = list();
    const follow = new FollowEnd(element);
    userScrollsTo(100);
    grow(500);
    follow.scroll();
    expect(element.scrollTop).toBe(100);
    userScrollsTo(1500 - 300 - 10);
    grow(100);
    follow.scroll();
    expect(element.scrollTop).toBe(1600);
  });

  it("jumps to the end when forced, for example after the person sends a message", () => {
    const { element, grow, userScrollsTo } = list();
    const follow = new FollowEnd(element);
    userScrollsTo(100);
    grow(500);
    follow.scroll(true);
    expect(element.scrollTop).toBe(1500);
    grow(50);
    follow.scroll();
    expect(element.scrollTop).toBe(1550);
  });

  it("tells when it stops and starts following, so a button can show and hide", () => {
    const { element, grow, userScrollsTo } = list();
    const changes: boolean[] = [];
    const follow = new FollowEnd(element, { onChange: (following) => changes.push(following) });
    follow.scroll();
    userScrollsTo(100);
    userScrollsTo(120);
    grow(500);
    follow.scroll();
    expect(changes).toEqual([false]);
    follow.scroll(true);
    expect(changes).toEqual([false, true]);
    userScrollsTo(1500 - 300);
    expect(changes).toEqual([false, true]);
  });
});

describe("steadyVisibility", () => {
  it("shows the button only after the person has stayed away from the end, and hides it at once", () => {
    vi.useFakeTimers();
    const seen: boolean[] = [];
    const change = steadyVisibility((visible) => seen.push(visible), 250);
    change(false);
    vi.advanceTimersByTime(100);
    change(true); // back at the end within the delay: nothing was ever shown
    vi.advanceTimersByTime(500);
    expect(seen).toEqual([false]);
    seen.length = 0;
    change(false);
    vi.advanceTimersByTime(250);
    expect(seen).toEqual([true]);
    change(true);
    expect(seen).toEqual([true, false]);
    vi.useRealTimers();
  });
});
