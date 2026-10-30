// @vitest-environment jsdom
import fs from "node:fs";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { setupTestingPanel } from "../addin/taskpane/ui/testing-panel.js";
import type { TestingOverview } from "../addin/taskpane/bridge-client.js";

const overview = (extra: Partial<TestingOverview> = {}): TestingOverview => ({
  signIns: [{ name: "admin", origins: ["https://staging.example.com"], savedAt: "2026-10-29T10:00:00.000Z", url: "https://staging.example.com" }],
  notes: [{ site: "staging.example.com", notes: [{ name: "login", kind: "login", summary: "Use <b>the</b> saved session", updated: "2026-10-29" }] }],
  scripts: [{ site: "staging.example.com", scripts: [{ caseId: "TC-001", title: "Create a ticket", outcome: "fail_app_defect", tested: "2026-10-29" }] }],
  pendingSignIn: null,
  folders: { sessions: "S", knowledge: "K", scripts: "X" },
  ...extra,
});

function setup(data = overview()) {
  document.body.innerHTML = fs.readFileSync("addin/taskpane/index.html", "utf8").match(/<body[^>]*>([\s\S]*)<\/body>/)![1]!;
  const client = {
    testingOverview: vi.fn(async () => data),
    testingView: vi.fn(async () => "// the script text"),
    testingAction: vi.fn(async (body: Record<string, unknown>) => body.action === "sign-in-finish" ? { name: "tester", origins: ["https://a.test"] } : {}),
    testingDelete: vi.fn(async () => undefined),
  };
  const panel = setupTestingPanel(() => client as never);
  const el = (id: string) => document.getElementById(id) as HTMLElement;
  const buttons = (root: HTMLElement) => Array.from(root.querySelectorAll("button"));
  const clickText = (root: HTMLElement, text: string) => buttons(root).find(b => b.textContent === text)!.click();
  return { client, panel, el, buttons, clickText };
}
const flush = () => new Promise(resolve => setTimeout(resolve, 0));

describe("the testing panel", () => {
  let ctx: ReturnType<typeof setup>;
  beforeEach(() => { ctx = setup(); });

  it("lists sign-ins, notes and scripts, showing text as text", async () => {
    await ctx.panel.refresh();
    expect(ctx.el("testing-signins").textContent).toContain("admin");
    expect(ctx.el("testing-signins").textContent).toContain("https://staging.example.com");
    expect(ctx.el("testing-notes").textContent).toContain("Use <b>the</b> saved session");
    expect(ctx.el("testing-notes").querySelector("b")).toBeNull();
    expect(ctx.el("testing-scripts").textContent).toContain("application defect");
  });

  it("needs a second click to delete, and then removes and reloads", async () => {
    await ctx.panel.refresh();
    ctx.clickText(ctx.el("testing-scripts"), "Delete");
    await flush();
    expect(ctx.client.testingDelete).not.toHaveBeenCalled();
    ctx.clickText(ctx.el("testing-scripts"), "Click again to delete");
    await flush();
    expect(ctx.client.testingDelete).toHaveBeenCalledWith({ kind: "script", site: "staging.example.com", caseId: "TC-001" });
    expect(ctx.client.testingOverview).toHaveBeenCalledTimes(2);
  });

  it("shows and hides a script's text", async () => {
    await ctx.panel.refresh();
    ctx.clickText(ctx.el("testing-scripts"), "View");
    await flush();
    expect(ctx.el("testing-scripts").querySelector("pre")!.textContent).toBe("// the script text");
    expect(ctx.client.testingView).toHaveBeenCalledWith("script", "staging.example.com", "TC-001");
    ctx.clickText(ctx.el("testing-scripts"), "Hide");
    expect(ctx.el("testing-scripts").querySelector("pre")).toBeNull();
  });

  it("opens a browser for a new sign-in, waits for 'I'm signed in', then saves it", async () => {
    await ctx.panel.refresh();
    (document.getElementById("testing-signin-name") as HTMLInputElement).value = "Tester";
    (document.getElementById("testing-signin-url") as HTMLInputElement).value = "https://a.test";
    ctx.el("testing-signin-form").dispatchEvent(new Event("submit", { cancelable: true }));
    await flush();
    expect(ctx.client.testingAction).toHaveBeenCalledWith({ action: "sign-in-start", name: "tester", url: "https://a.test" });
    expect(ctx.el("testing-signin-wait").hidden).toBe(false);
    expect(ctx.el("testing-signin-form").hidden).toBe(true);
    ctx.el("testing-signin-done").click();
    await flush();
    expect(ctx.client.testingAction).toHaveBeenLastCalledWith({ action: "sign-in-finish" });
    expect(ctx.el("testing-status").textContent).toContain("Saved sign-in 'tester'");
  });

  it("shows a window that is already open when the panel is reopened", async () => {
    const reopened = setup(overview({ pendingSignIn: { name: "admin", url: "https://staging.example.com" } }));
    await reopened.panel.refresh();
    expect(reopened.el("testing-signin-wait").hidden).toBe(false);
    expect(reopened.el("testing-signin-wait-text").textContent).toContain("save it as 'admin'");
  });
});
