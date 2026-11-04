import { describe, expect, it, vi } from "vitest";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { ExternalTools, type BrowserFrame } from "../bridge/src/external-tools.js";
import { describeBrowserAction, displayUrl, extractImage, parsePageInfo } from "../bridge/src/browser-frames.js";
import { ServerMessageSchema } from "../protocol/src/index.js";

describe("browser action labels", () => {
  it("say what happened without repeating anything typed", () => {
    expect(describeBrowserAction("browser_navigate", { url: "https://staging.example.com/web?token=SECRET#frag" })).toBe("Opening staging.example.com/web");
    expect(describeBrowserAction("browser_click", { element: "Confirm button" })).toBe('Clicking "Confirm button"');
    const typed = describeBrowserAction("browser_type", { element: "Password box", text: "hunter2" });
    expect(typed).toBe('Typing into "Password box"');
    expect(typed).not.toContain("hunter2");
    expect(describeBrowserAction("browser_fill_form", { fields: [{ name: "a", value: "secret" }, { name: "b", value: "x" }] })).toBe("Filling 2 field(s)");
    expect(describeBrowserAction("browser_evaluate", { function: "() => document.cookie" })).toBe("Running a script on the page");
  });

  it("show named keys only, because a single key could be part of a password", () => {
    expect(describeBrowserAction("browser_press_key", { key: "Enter" })).toBe("Pressing Enter");
    expect(describeBrowserAction("browser_press_key", { key: "Control+ArrowDown" })).toBe("Pressing Control+ArrowDown");
    expect(describeBrowserAction("browser_press_key", { key: "p" })).toBe("Pressing a key");
    expect(describeBrowserAction("browser_press_key", { key: "hunter2" })).toBe("Pressing a key");
  });

  it("shortens long element descriptions and survives odd input", () => {
    expect(describeBrowserAction("browser_click", { element: "x".repeat(200) }).length).toBeLessThan(80);
    expect(describeBrowserAction("browser_click", {})).toBe("Clicking an element");
    expect(describeBrowserAction("browser_something_new")).toBe("something new");
  });
});

describe("reading Playwright MCP output", () => {
  it("never shows a query string or fragment", () => {
    expect(displayUrl("https://a.test/web/login?db=x&token=abc#section")).toBe("a.test/web/login");
    expect(displayUrl("https://a.test/")).toBe("a.test");
    expect(displayUrl("about:blank")).toBe("about:blank");
    expect(displayUrl("not a url")).toBeUndefined();
    expect(displayUrl(undefined)).toBeUndefined();
  });

  it("finds the page address and title, and the picture", () => {
    const output = { content: [{ type: "text", text: "### Page\n- Page URL: https://a.test/web?x=1\n- Page Title: Home | Odoo\n### Snapshot" }, { type: "image", mimeType: "image/jpeg", data: "QUJD" }] };
    expect(parsePageInfo(output)).toEqual({ url: "a.test/web", title: "Home | Odoo" });
    expect(extractImage(output)).toEqual({ mimeType: "image/jpeg", data: "QUJD" });
    expect(parsePageInfo({ content: [{ type: "text", text: "nothing" }] })).toEqual({ url: undefined, title: undefined });
    expect(extractImage({ content: [{ type: "image", mimeType: "image/gif", data: "x" }] })).toBeNull();
    expect(extractImage(undefined)).toBeNull();
  });
});

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
const text = (body: string) => ({ content: [{ type: "text", text: body }] });
const SHOT = { content: [{ type: "text", text: "### Result" }, { type: "image", mimeType: "image/jpeg", data: "SHOT" }] };

function withFakeBrowser(responses: Record<string, unknown> = {}) {
  const tools = new ExternalTools();
  const calls: string[] = [];
  let active = 0, maxActive = 0;
  const client = {
    listTools: async () => ({ tools: ["browser_navigate", "browser_click", "browser_type", "browser_snapshot", "browser_take_screenshot", "browser_console_messages"].map(name => ({ name })) }),
    callTool: async ({ name }: { name: string }) => {
      active++; maxActive = Math.max(maxActive, active); calls.push(name);
      await sleep(5);
      active--;
      const response = responses[name];
      if (response instanceof Error) throw response;
      return response ?? text(`### Page\n- Page URL: https://a.test/web?token=abc#x\n- Page Title: Home`);
    },
  };
  vi.spyOn(tools as never, "client" as never).mockResolvedValue(client as never);
  const frames: BrowserFrame[] = [];
  const sink = (frame: BrowserFrame) => { frames.push(frame); };
  return { tools, calls, frames, sink, maxActive: () => maxActive };
}

describe("live view frames", () => {
  it("photographs the page after an action, with the address stripped of its query", async () => {
    const { tools, frames, sink, calls } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await tools.browser("s", "browser_navigate", { url: "https://a.test/web?token=abc" }, undefined, sink);
    await tools.browser("s", "browser_console_messages", {}); // waits for the queued picture
    expect(calls.filter(name => name === "browser_take_screenshot")).toHaveLength(1);
    expect(frames).toEqual([{ seq: 0, label: "Opening a.test/web", url: "a.test/web", title: "Home", mimeType: "image/jpeg", data: "SHOT" }]);
  });

  it("returns the assistant's result unchanged and does not wait for the picture", async () => {
    const original = text("### Page\n- Page URL: https://a.test/\n");
    const { tools, sink } = withFakeBrowser({ browser_navigate: original, browser_take_screenshot: SHOT });
    expect(await tools.browser("s", "browser_navigate", { url: "https://a.test/" }, undefined, sink)).toBe(original);
  });

  it("runs the assistant's calls one at a time, including the picture", async () => {
    const { tools, maxActive, sink } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await Promise.all([
      tools.browser("s", "browser_navigate", { url: "https://a.test/" }, undefined, sink),
      tools.browser("s", "browser_click", { element: "A" }, undefined, sink),
      tools.browser("s", "browser_snapshot", {}, undefined, sink),
    ]);
    await tools.browser("s", "browser_console_messages", {});
    expect(maxActive()).toBe(1);
  });

  it("merges a burst of actions into one picture labelled by the last of them", async () => {
    const { tools, frames, sink, calls } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await Promise.all([
      tools.browser("s", "browser_click", { element: "First" }, undefined, sink),
      tools.browser("s", "browser_click", { element: "Second" }, undefined, sink),
      tools.browser("s", "browser_click", { element: "Third" }, undefined, sink),
    ]);
    await tools.browser("s", "browser_console_messages", {});
    expect(calls.filter(name => name === "browser_take_screenshot")).toHaveLength(1);
    expect(frames.map(frame => frame.label)).toEqual(['Clicking "Third"']);
  });

  it("uses the assistant's own screenshot instead of taking another", async () => {
    const { tools, frames, sink, calls } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await tools.browser("s", "browser_take_screenshot", {}, undefined, sink);
    await tools.browser("s", "browser_console_messages", {});
    expect(calls.filter(name => name === "browser_take_screenshot")).toHaveLength(1);
    expect(frames).toHaveLength(1);
    expect(frames[0]!.label).toBe("Taking a screenshot");
  });

  it("shows the first read of a page, then stays quiet for further reads", async () => {
    const { tools, frames, sink } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await tools.browser("s", "browser_snapshot", {}, undefined, sink);
    await tools.browser("s", "browser_snapshot", {}, undefined, sink);
    await tools.browser("s", "browser_console_messages", {}, undefined, sink);
    expect(frames).toHaveLength(1);
  });

  it("numbers pictures per session and keeps sessions apart", async () => {
    const { tools, frames, sink } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await tools.browser("a", "browser_click", { element: "X" }, undefined, sink);
    await tools.browser("a", "browser_click", { element: "Y" }, undefined, sink);
    await tools.browser("a", "browser_console_messages", {});
    await tools.browser("b", "browser_click", { element: "Z" }, undefined, sink);
    await tools.browser("b", "browser_console_messages", {});
    expect(frames.map(frame => frame.seq)).toEqual([0, 1, 0]);
  });

  it("marks a failed action as failed (using the picture taken for the assistant), and shows nothing for failed screenshots", async () => {
    const failed = withFakeBrowser({ browser_click: { isError: true, content: [{ type: "text", text: "### Error" }] }, browser_take_screenshot: SHOT });
    await failed.tools.browser("s", "browser_click", { element: "X" }, undefined, failed.sink);
    await failed.tools.browser("s", "browser_console_messages", {});
    expect(failed.frames.map(frame => frame.label)).toEqual(['Failed: Clicking "X"']);
    expect(failed.calls.filter(name => name === "browser_take_screenshot")).toHaveLength(1); // one picture serves both the assistant and the person

    const broken = withFakeBrowser({ browser_take_screenshot: new Error("boom") });
    await expect(broken.tools.browser("s", "browser_click", { element: "X" }, undefined, broken.sink)).resolves.toBeDefined();
    await broken.tools.browser("s", "browser_console_messages", {});
    expect(broken.frames).toEqual([]);

    const dialog = withFakeBrowser({ browser_take_screenshot: { isError: true, content: [{ type: "text", text: "### Error\nTool does not handle the modal state." }] } });
    await dialog.tools.browser("s", "browser_click", { element: "X" }, undefined, dialog.sink);
    await dialog.tools.browser("s", "browser_console_messages", {});
    expect(dialog.frames).toEqual([]);
  });

  it("skips pictures that are too large to send, and survives a sink that throws", async () => {
    const huge = withFakeBrowser({ browser_take_screenshot: { content: [{ type: "image", mimeType: "image/jpeg", data: "A".repeat(2_600_000) }] } });
    await huge.tools.browser("s", "browser_click", { element: "X" }, undefined, huge.sink);
    await huge.tools.browser("s", "browser_console_messages", {});
    expect(huge.frames).toEqual([]);

    const { tools } = withFakeBrowser({ browser_take_screenshot: SHOT });
    const throwing = () => { throw new Error("pane closed"); };
    await expect(tools.browser("s", "browser_click", { element: "X" }, undefined, throwing)).resolves.toBeDefined();
    await expect(tools.browser("s", "browser_console_messages", {})).resolves.toBeDefined();
  });

  it("takes no pictures unless asked to", async () => {
    const { tools, calls } = withFakeBrowser({ browser_take_screenshot: SHOT });
    await tools.browser("s", "browser_click", { element: "X" });
    await tools.browser("s", "browser_console_messages", {});
    expect(calls).not.toContain("browser_take_screenshot");
  });
});

describe("browser_frame messages", () => {
  const frame = { type: "browser_frame", requestId: "r1", seq: 0, label: "Opening a.test", mimeType: "image/jpeg", data: "QUJD" };

  it("are part of the protocol and reject oversized or unexpected pictures", () => {
    expect(ServerMessageSchema.safeParse(frame).success).toBe(true);
    expect(ServerMessageSchema.safeParse({ ...frame, mimeType: "image/svg+xml" }).success).toBe(false);
    expect(ServerMessageSchema.safeParse({ ...frame, data: "A".repeat(2_500_001) }).success).toBe(false);
    expect(ServerMessageSchema.safeParse({ ...frame, data: "" }).success).toBe(false);
  });

  it("are sent to the pane for the active request and dropped once it has ended", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("frames", null);
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000);
    session.capabilities.playwright = true;
    Object.defineProperty(session, "isConnected", { get: () => true });
    vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: true } });
    const sent: unknown[] = [];
    vi.spyOn(session, "send").mockImplementation(message => { sent.push(message); return true; });
    const picture = { seq: 0, label: "Opening a.test", mimeType: "image/jpeg" as const, data: "QUJD" };
    vi.spyOn(gateway.externalTools, "browser").mockImplementation(async (_id, _tool, _args, _signal, show) => {
      show?.(picture);
      session.activeRequestId = "a-newer-request";
      show?.({ ...picture, seq: 1 });
      return { tools: [] };
    });
    await gateway.invoke({ sessionId: session.id, tool: "playwright_browser", args: { tool: "browser_navigate", arguments: { url: "https://a.test/" } } });
    expect(sent).toEqual([{ type: "browser_frame", requestId: "no-turn", ...picture }]);
  });
});
