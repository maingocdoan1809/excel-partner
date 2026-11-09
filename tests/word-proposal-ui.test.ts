// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";
import { showProposalReview, diffParts } from "../addin/taskpane/ui/proposal.js";
import type { ProposalRequest } from "../addin/taskpane/executor/word/index.js";

afterEach(() => document.body.replaceChildren());

const request: ProposalRequest = {
  title: "Tone edits",
  items: [
    { index: 0, oldText: "grew 5% quickly", newText: "rose 5% quickly", reason: "plainer", alternatives: ["increased 5% quickly"], found: true, ambiguous: false },
    { index: 1, oldText: "Revenue", newText: "Sales", found: true, ambiguous: true },
    { index: 2, oldText: "missing", newText: "x", found: false, ambiguous: false },
  ],
};
const boxes = () => [...document.querySelectorAll<HTMLInputElement>('.proposal-item-head input[type="checkbox"]')];
const submit = () => document.querySelector("form")!.dispatchEvent(new Event("submit", { cancelable: true }));

describe("proposal review card", () => {
  it("checks found, unambiguous edits only and disables the ones not found", () => {
    void showProposalReview(document.body, request, new AbortController().signal);
    expect(boxes().map((b) => [b.checked, b.disabled])).toEqual([[true, false], [false, false], [false, true]]);
    expect(document.querySelector(".proposal-title")?.textContent).toBe("Tone edits");
    expect(document.querySelectorAll(".proposal-badge")).toHaveLength(2);
    expect(document.querySelector('button[type="submit"]')?.textContent).toBe("Apply selected (1)");
  });

  it("shows a word-level before / after diff with textContent only", () => {
    void showProposalReview(document.body, request, new AbortController().signal);
    const item = document.querySelector(".proposal-item")!;
    expect(item.querySelector(".proposal-before")?.textContent).toBe("grew 5% quickly");
    expect(item.querySelector(".proposal-del")?.textContent).toBe("grew");
    expect(item.querySelector(".proposal-ins")?.textContent).toBe("rose");
  });

  it("never turns document text into markup", () => {
    const hostile = { index: 0, oldText: '<img src=x onerror="alert(1)">', newText: "<b>bold</b>", reason: "<script>x()</script>", found: true, ambiguous: false };
    void showProposalReview(document.body, { items: [hostile] }, new AbortController().signal);
    expect(document.querySelector("img, b, script")).toBeNull();
    expect(document.querySelector(".proposal-before")?.textContent).toBe('<img src=x onerror="alert(1)">');
    expect(document.querySelector(".proposal-reason")?.textContent).toBe("<script>x()</script>");
  });

  it("returns the chosen wording for the checked edits on Apply", async () => {
    const pending = showProposalReview(document.body, request, new AbortController().signal);
    const radios = document.querySelectorAll<HTMLInputElement>('.proposal-choice input[type="radio"]');
    expect(radios).toHaveLength(2);
    radios[1]!.checked = true;
    radios[1]!.dispatchEvent(new Event("change"));
    expect(document.querySelector(".proposal-ins")?.textContent).toBe("increased");
    boxes()[1]!.checked = true;
    boxes()[1]!.dispatchEvent(new Event("change"));
    submit();
    expect(await pending).toEqual({ cancelled: false, accepted: [{ index: 0, text: "increased 5% quickly" }, { index: 1, text: "Sales" }] });
    expect(document.querySelector("form")).toBeNull();
  });

  it("disables Apply when nothing is checked, and Reject all accepts nothing without cancelling", async () => {
    const pending = showProposalReview(document.body, request, new AbortController().signal);
    boxes()[0]!.checked = false;
    boxes()[0]!.dispatchEvent(new Event("change"));
    expect(document.querySelector<HTMLButtonElement>('button[type="submit"]')!.disabled).toBe(true);
    [...document.querySelectorAll("button")].find((b) => b.textContent === "Reject all")!.click();
    expect(await pending).toEqual({ cancelled: false, accepted: [] });
  });

  it("cancels on Escape", async () => {
    const pending = showProposalReview(document.body, request, new AbortController().signal);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(await pending).toEqual({ cancelled: true, accepted: [] });
    expect(document.querySelector("form")).toBeNull();
  });

  it("cancels when the signal aborts, and not at all when it is already aborted", async () => {
    const controller = new AbortController();
    const pending = showProposalReview(document.body, request, controller.signal);
    controller.abort();
    expect(await pending).toEqual({ cancelled: true, accepted: [] });
    expect(document.querySelector("form")).toBeNull();
    controller.abort();
    expect(await showProposalReview(document.body, request, controller.signal)).toEqual({ cancelled: true, accepted: [] });
    expect(document.querySelector("form")).toBeNull();
  });

  it("diffParts falls back to plain text for very long edits", () => {
    const long = "word ".repeat(2000);
    expect(diffParts(long, `${long}x`).before).toEqual([{ text: long, changed: false }]);
  });
});
