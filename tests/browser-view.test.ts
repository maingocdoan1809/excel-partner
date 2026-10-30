// @vitest-environment jsdom
import fs from "node:fs";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { BrowserView, MAX_BROWSER_FRAMES } from "../addin/taskpane/ui/browser-view.js";
import type { BrowserFrameMessage } from "../protocol/src/index.js";

const frame = (n: number, extra: Partial<BrowserFrameMessage> = {}): BrowserFrameMessage =>
  ({ type: "browser_frame", requestId: "r1", seq: n, label: `Step ${n}`, mimeType: "image/jpeg", data: `IMG${n}`, url: "a.test/web", ...extra });

function setup() {
  document.body.innerHTML = fs.readFileSync("addin/taskpane/index.html", "utf8").match(/<body[^>]*>([\s\S]*)<\/body>/)![1]!;
  const el = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const open = vi.fn();
  const view = new BrowserView({
    dock: el("browser-dock"), dockLabel: el("browser-dock-label"), dockStep: el("browser-dock-step"), dockUrl: el("browser-dock-url"),
    dockImage: el<HTMLImageElement>("browser-dock-image"), dockOpen: el<HTMLButtonElement>("browser-dock-open"),
    dockExpand: el<HTMLButtonElement>("browser-dock-expand"), dockClose: el<HTMLButtonElement>("browser-dock-close"),
    panelLabel: el("browser-panel-label"), panelStep: el("browser-panel-step"), panelUrl: el("browser-panel-url"),
    panelImage: el<HTMLImageElement>("browser-panel-image"), panelPrev: el<HTMLButtonElement>("browser-prev"), panelNext: el<HTMLButtonElement>("browser-next"),
  }, open);
  return { view, open, el };
}

describe("the assistant's browser view", () => {
  let ctx: ReturnType<typeof setup>;
  beforeEach(() => { ctx = setup(); });

  it("is hidden until the first picture arrives", () => {
    expect(ctx.el("browser-dock").hidden).toBe(true);
    expect(ctx.el("panel-browser").hidden).toBe(true);
    ctx.view.push(frame(0, { label: 'Opening a.test/web', title: "Home" }));
    expect(ctx.el("browser-dock").hidden).toBe(false);
    expect(ctx.el("browser-dock-label").textContent).toBe("Opening a.test/web");
    expect(ctx.el("browser-dock-url").textContent).toBe("Home · a.test/web");
    expect(ctx.el("browser-dock-step").textContent).toBe("Step 1");
    expect(ctx.el<HTMLImageElement>("browser-dock-image").src).toBe("data:image/jpeg;base64,IMG0");
    expect(ctx.el<HTMLImageElement>("browser-dock-image").alt).toBe("Browser after: Opening a.test/web");
  });

  it("switches the large picture between fit-to-view and actual size by clicking it", () => {
    ctx.view.push(frame(0));
    const holder = ctx.el("browser-panel-image").parentElement!;
    expect(holder.classList.contains("actual")).toBe(false);
    ctx.el("browser-panel-image").click();
    expect(holder.classList.contains("actual")).toBe(true);
    ctx.el("browser-panel-image").click();
    expect(holder.classList.contains("actual")).toBe(false);
  });

  it("reports how many pictures there are so the menu can offer a way back after the small view is closed", () => {
    const counts: number[] = [];
    const again = new BrowserView({
      dock: ctx.el("browser-dock"), dockLabel: ctx.el("browser-dock-label"), dockStep: ctx.el("browser-dock-step"), dockUrl: ctx.el("browser-dock-url"),
      dockImage: ctx.el<HTMLImageElement>("browser-dock-image"), dockOpen: ctx.el<HTMLButtonElement>("browser-dock-open"),
      dockExpand: ctx.el<HTMLButtonElement>("browser-dock-expand"), dockClose: ctx.el<HTMLButtonElement>("browser-dock-close"),
      panelLabel: ctx.el("browser-panel-label"), panelStep: ctx.el("browser-panel-step"), panelUrl: ctx.el("browser-panel-url"),
      panelImage: ctx.el<HTMLImageElement>("browser-panel-image"), panelPrev: ctx.el<HTMLButtonElement>("browser-prev"), panelNext: ctx.el<HTMLButtonElement>("browser-next"),
    }, () => undefined, count => counts.push(count));
    again.push(frame(0));
    again.push(frame(1));
    ctx.el("browser-dock-close").click();
    expect(again.count).toBe(2);
    again.reset();
    expect(counts.at(-1)).toBe(0);
    expect(counts).toContain(2);
    expect(ctx.el("menu-browser")).not.toBeNull();
  });

  it("opens the larger view from the picture and from the expand button", () => {
    ctx.view.push(frame(0));
    ctx.el("browser-dock-open").click();
    ctx.el("browser-dock-expand").click();
    expect(ctx.open).toHaveBeenCalledTimes(2);
  });

  it("can be hidden, and comes back with the next picture", () => {
    ctx.view.push(frame(0));
    ctx.el("browser-dock-close").click();
    expect(ctx.el("browser-dock").hidden).toBe(true);
    ctx.view.push(frame(1));
    expect(ctx.el("browser-dock").hidden).toBe(false);
  });

  it("brings the small view back above the message box on request, and only when there is a picture", () => {
    ctx.view.showDock();
    expect(ctx.el("browser-dock").hidden).toBe(true);
    ctx.view.push(frame(0));
    ctx.el("browser-dock-close").click();
    expect(ctx.el("browser-dock").hidden).toBe(true);
    ctx.view.showDock();
    expect(ctx.el("browser-dock").hidden).toBe(false);
    expect(ctx.el<HTMLImageElement>("browser-dock-image").src).toBe("data:image/jpeg;base64,IMG0");
  });

  it("follows the newest picture until the person steps back, then stays where they are", () => {
    for (let n = 0; n < 3; n++) ctx.view.push(frame(n));
    expect(ctx.el("browser-panel-step").textContent).toBe("3 / 3");
    expect(ctx.el<HTMLButtonElement>("browser-next").disabled).toBe(true);
    ctx.el("browser-prev").click();
    expect(ctx.el("browser-panel-step").textContent).toBe("2 / 3");
    expect(ctx.el<HTMLImageElement>("browser-panel-image").src).toContain("IMG1");
    ctx.view.push(frame(3));
    expect(ctx.el("browser-panel-step").textContent).toBe("2 / 4");
    expect(ctx.el<HTMLImageElement>("browser-panel-image").src).toContain("IMG1");
    expect(ctx.el<HTMLImageElement>("browser-dock-image").src).toContain("IMG3"); // the small view always shows the newest
    ctx.el("browser-next").click();
    ctx.el("browser-next").click();
    expect(ctx.el("browser-panel-step").textContent).toBe("4 / 4");
    ctx.view.push(frame(4));
    expect(ctx.el("browser-panel-step").textContent).toBe("5 / 5"); // back at the end, so it follows again
  });

  it("disables stepping at both ends", () => {
    ctx.view.push(frame(0));
    expect(ctx.el<HTMLButtonElement>("browser-prev").disabled).toBe(true);
    expect(ctx.el<HTMLButtonElement>("browser-next").disabled).toBe(true);
    ctx.view.push(frame(1));
    expect(ctx.el<HTMLButtonElement>("browser-prev").disabled).toBe(false);
  });

  it(`keeps only the last ${MAX_BROWSER_FRAMES} pictures`, () => {
    for (let n = 0; n < MAX_BROWSER_FRAMES + 5; n++) ctx.view.push(frame(n));
    expect(ctx.view.count).toBe(MAX_BROWSER_FRAMES);
    expect(ctx.el<HTMLImageElement>("browser-panel-image").src).toContain(`IMG${MAX_BROWSER_FRAMES + 4}`);
  });

  it("forgets everything on reset", () => {
    ctx.view.push(frame(0));
    ctx.view.reset();
    expect(ctx.view.count).toBe(0);
    expect(ctx.el("browser-dock").hidden).toBe(true);
    expect(ctx.el<HTMLImageElement>("browser-dock-image").getAttribute("src")).toBeNull();
    expect(ctx.el<HTMLImageElement>("browser-panel-image").getAttribute("src")).toBeNull();
    expect(ctx.el("browser-panel-label").textContent).toContain("Nothing yet");
  });

  it("shows picture text as text, never as markup", () => {
    ctx.view.push(frame(0, { label: '<img src=x onerror=alert(1)>', title: "<b>t</b>" }));
    expect(ctx.el("browser-dock-label").querySelector("img")).toBeNull();
    expect(ctx.el("browser-dock-label").textContent).toBe('<img src=x onerror=alert(1)>');
    expect(ctx.el("browser-dock-url").querySelector("b")).toBeNull();
  });

  it("is part of the page the task pane loads", () => {
    const html = fs.readFileSync("addin/taskpane/index.html", "utf8");
    for (const id of ["browser-dock", "browser-dock-image", "panel-browser", "browser-panel-image", "browser-prev", "browser-next"]) expect(html).toContain(`id="${id}"`);
    expect(html).toContain('data-close="browser"');
  });
});
