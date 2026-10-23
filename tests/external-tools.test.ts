import { describe, expect, it, vi } from "vitest";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { exposedToolNames, parseToolInput, supportsNativeWebSearch } from "../protocol/src/index.js";

function fixture() {
  const sessions = new SessionManager({ ttlMs: 60_000 });
  const session = sessions.create("external-tools", null);
  const gateway = new ToolGateway(sessions, { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 }, 1000);
  const call = (tool: string, args: unknown) => gateway.invoke({ sessionId: session.id, tool, args });
  return { session, gateway, call };
}

describe("optional external tools", () => {
  it("only exposes opted-in Playwright and never adds a search MCP", () => {
    expect((["claude", "codex", "copilot"] as const).every(supportsNativeWebSearch)).toBe(true);
    expect(supportsNativeWebSearch("opencode")).toBe(false);
    expect(exposedToolNames(true)).not.toContain("web_search");
    expect(exposedToolNames(true)).not.toContain("playwright_browser");
    expect(exposedToolNames(false, { webSearch: true })).not.toContain("web_search");
    expect(exposedToolNames(false, { webSearch: true })).not.toContain("playwright_browser");
    expect(exposedToolNames(false, { playwright: true })).toContain("playwright_browser");
    expect(() => parseToolInput("playwright_browser", { tool: "shell" })).toThrow();
  });

  it("blocks direct calls while disabled, without contacting external services", async () => {
    const { gateway, call } = fixture();
    const browser = vi.spyOn(gateway.externalTools, "browser");
    expect(await call("web_search", { query: "test" })).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
    expect(await call("playwright_browser", { tool: "list_tools" })).toMatchObject({ ok: false, error: { code: "unsupported" } });
    expect(browser).not.toHaveBeenCalled();
  });

  it("requires browser approval and performs the exact approved action", async () => {
    const { session, gateway, call } = fixture();
    session.capabilities.playwright = true;
    Object.defineProperty(session, "isConnected", { get: () => true });
    const request = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: false } });
    const browser = vi.spyOn(gateway.externalTools, "browser").mockResolvedValue({ tools: [] });
    const args = { tool: "browser_navigate", arguments: { url: "https://example.com" } };
    expect(await call("playwright_browser", args)).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(browser).not.toHaveBeenCalled();
    request.mockResolvedValue({ ok: true, data: { approved: true } });
    expect(await call("playwright_browser", args)).toEqual({ ok: true, data: { tools: [] } });
    expect(request).toHaveBeenLastCalledWith("__authorize_browser", args, 45_000);
    expect(browser).toHaveBeenCalledWith(session.id, args.tool, args.arguments, undefined);
  });

  it("does not start the browser if Stop arrives during approval", async () => {
    const { session, gateway, call } = fixture();
    session.capabilities.playwright = true;
    session.turnController = new AbortController();
    Object.defineProperty(session, "isConnected", { get: () => true });
    vi.spyOn(session, "requestTool").mockImplementation(async () => {
      session.turnController!.abort();
      return { ok: true, data: { approved: true } };
    });
    const browser = vi.spyOn(gateway.externalTools, "browser");
    expect(await call("playwright_browser", { tool: "list_tools" })).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(browser).not.toHaveBeenCalled();
  });
});
