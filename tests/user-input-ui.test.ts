// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";
import { UserInputRequestSchema } from "../protocol/src/index.js";
import { showUserInput } from "../addin/taskpane/ui/user-input.js";

afterEach(() => document.body.replaceChildren());
const request = (multiSelect = false) => UserInputRequestSchema.parse({ questions: [{ id: "q", question: "Where?",
  options: [{ label: "Existing", description: "Keep it here" }, { label: "New" }], multiSelect }] });
const submit = () => document.querySelector("form")!.dispatchEvent(new Event("submit", { cancelable: true }));

describe("native question UI", () => {
  it("requires an explicit answer and submits the chosen option", async () => {
    const pending = showUserInput(document.body, request(), new AbortController().signal);
    expect(document.querySelector<HTMLInputElement>("input:checked")).toBeNull();
    submit();
    expect(document.querySelector('[role="alert"]')?.textContent).toContain("answer every question");
    document.querySelector<HTMLInputElement>('input[value="New"]')!.checked = true;
    submit();
    expect(await pending).toEqual({ cancelled: false, answers: { q: ["New"] } });
    expect(document.querySelector("form")).toBeNull();
  });
  it("supports multiple choices and a custom answer", async () => {
    const pending = showUserInput(document.body, request(true), new AbortController().signal);
    document.querySelector<HTMLInputElement>('input[value="Existing"]')!.checked = true;
    document.querySelector<HTMLInputElement>('input[type="text"]')!.value = "Another workbook";
    submit();
    expect(await pending).toMatchObject({ answers: { q: ["Existing", "Another workbook"] } });
  });
  it("switches cleanly between single-choice and freeform answers", async () => {
    const pending = showUserInput(document.body, request(), new AbortController().signal);
    const option = document.querySelector<HTMLInputElement>('input[value="Existing"]')!;
    option.checked = true;
    const input = document.querySelector<HTMLInputElement>('input[type="text"]')!;
    input.value = "Custom";
    input.dispatchEvent(new Event("input"));
    expect(option.checked).toBe(false);
    submit();
    expect(await pending).toMatchObject({ answers: { q: ["Custom"] } });
  });
  it.each(["skip", "abort"])("removes the card on %s without choosing an answer", async action => {
    const controller = new AbortController();
    const pending = showUserInput(document.body, request(), controller.signal);
    if (action === "abort") controller.abort();
    else document.querySelector<HTMLButtonElement>('button[type="button"]')!.click();
    expect(await pending).toEqual({ cancelled: true, answers: {} });
    expect(document.querySelector("form")).toBeNull();
  });
  it("renders question content as text and secrets as password fields", async () => {
    const controller = new AbortController();
    const pending = showUserInput(document.body, UserInputRequestSchema.parse({ questions: [
      { id: "secret", question: '<img src=x onerror="alert(1)">', isSecret: true },
    ] }), controller.signal);
    expect(document.querySelector("img")).toBeNull();
    expect(document.querySelector('input[type="password"]')).not.toBeNull();
    controller.abort();
    await pending;
  });
});
