// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import fs from "node:fs";
import { setupConnectedFiles } from "../addin/taskpane/ui/connected-files.js";

beforeEach(() => {
  const html = fs.readFileSync("addin/taskpane/index.html", "utf8");
  document.body.innerHTML = html.split("<body>")[1]!.split("</body>")[0]!;
});
afterEach(() => document.body.replaceChildren());
const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

describe("Connected files button", () => {
  it("opens and closes the list from the button, outside clicks and Escape", () => {
    setupConnectedFiles();
    const button = element<HTMLButtonElement>("files-button");
    const popover = element("files-popover");
    expect(popover.hidden).toBe(true);
    button.click();
    expect(popover.hidden).toBe(false);
    expect(button.getAttribute("aria-expanded")).toBe("true");
    popover.click();
    expect(popover.hidden).toBe(false);
    document.body.click();
    expect(popover.hidden).toBe(true);
    button.click();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(popover.hidden).toBe(true);
    expect(button.getAttribute("aria-expanded")).toBe("false");
  });

  it("lists each file with its full name on hover, marks the current one, and truncates in CSS not in text", () => {
    const files = setupConnectedFiles();
    const long = "AstraLink_データ移行マッピング_最終版_v12_2026-11-05.xlsx";
    files.render([{ id: "f2", name: long, host: "excel", current: true }, { id: "f1", name: "Budget.xlsx", host: "excel", current: false }]);
    const rows = Array.from(document.querySelectorAll<HTMLElement>("#files-list .files-row"));
    expect(rows.map((row) => row.querySelector(".files-name")!.textContent)).toEqual([long, "Budget.xlsx"]);
    expect(rows[0]!.title).toBe(long);
    expect(rows[0]!.querySelector(".files-tag")?.textContent).toBe("This file");
    expect(rows[1]!.querySelector(".files-tag")).toBeNull();
    // It can also ask the other file's assistant to change that file, so the hint must not say the files are read-only.
    expect(element("files-hint").textContent).not.toContain("cannot change");
    expect(element("files-hint").textContent).toContain("make changes");
    expect(element("files-hint").textContent).toContain("approve each request");
  });

  it("shows how to connect another file when only this one is open, and a placeholder for an unnamed file", () => {
    const files = setupConnectedFiles();
    files.render([{ id: "f1", name: "  ", host: "excel", current: true }]);
    expect(document.querySelector("#files-list .files-name")!.textContent).toBe("Untitled workbook");
    expect(element("files-hint").textContent).toContain("Open another file");
  });

  it("shows which application each file is open in, with a placeholder name that fits the kind of file", () => {
    const files = setupConnectedFiles();
    files.render([
      { id: "f1", name: "Plan.docx", host: "word", current: true },
      { id: "f2", name: "", host: "powerpoint", current: false },
      { id: "f3", name: "Budget.xlsx", host: "excel", current: false },
    ]);
    const rows = Array.from(document.querySelectorAll<HTMLElement>("#files-list .files-row"));
    expect(rows.map((row) => row.dataset.host)).toEqual(["word", "powerpoint", "excel"]);
    expect(rows.map((row) => row.querySelector(".files-app")!.textContent)).toEqual(["Word", "PowerPoint", "Excel"]);
    expect(rows.map((row) => row.querySelector(".files-name")!.textContent)).toEqual(["Plan.docx", "Untitled presentation", "Budget.xlsx"]);
    expect(new Set(rows.map((row) => row.querySelector("svg")!.innerHTML)).size).toBe(3);
  });

  it("closes when a button that stops the click (like Settings) is pressed, and does not hold the Settings menu open", () => {
    setupConnectedFiles();
    const button = element<HTMLButtonElement>("files-button");
    const popover = element("files-popover");
    const other = document.createElement("button");
    other.addEventListener("click", (event) => event.stopPropagation());
    document.body.append(other);
    button.click();
    expect(popover.hidden).toBe(false);
    other.click();
    expect(popover.hidden).toBe(true);

    // Opening the files list must still reach the document, where the Settings menu closes itself.
    let reached = false;
    document.addEventListener("click", () => { reached = true; });
    button.click();
    expect(reached).toBe(true);
  });

  it("asks for current names each time the list is opened, and not when it closes", () => {
    let asked = 0;
    setupConnectedFiles({ onOpen: () => { asked++; } });
    const button = element<HTMLButtonElement>("files-button");
    button.click();
    expect(asked).toBe(1);
    button.click();
    expect(asked).toBe(1);
    button.click();
    expect(asked).toBe(2);
  });

  it("shows the green dot only while the bridge is connected", () => {
    const files = setupConnectedFiles();
    const button = element<HTMLButtonElement>("files-button");
    expect(button.dataset.connected).toBe("false");
    files.setConnected(true);
    expect(button.dataset.connected).toBe("true");
    files.setConnected(false);
    expect(button.dataset.connected).toBe("false");
  });
});
