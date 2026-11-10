// @vitest-environment jsdom
import fs from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { setupSystemPanel, type SystemInfo } from "../addin/taskpane/ui/system-panel.js";

const el = (id: string) => document.getElementById(id)!;
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
afterEach(() => { vi.useRealTimers(); vi.restoreAllMocks(); });
function fixture(busy = false) {
  document.body.innerHTML = fs.readFileSync("addin/taskpane/index.html", "utf8").match(/<body[^>]*>([\s\S]*)<\/body>/)![1]!;
  const info: SystemInfo = { version: "1.1.0", previous: "1.0.0", held: false, phase: "committed", cached: ["1.0.0"],
    providers: [{ provider: "copilot", selected: true, mode: "managed", version: "1.0.95", pendingVersion: null, available: true },
      { provider: "codex", selected: false, mode: "external", version: null, pendingVersion: "0.162.1", available: null }] };
  const client = { systemInfo: vi.fn(async () => info), systemAction: vi.fn(async (_action: string, _provider?: string, _version?: string) => {}),
    systemDiagnostics: vi.fn(async () => ({ diagnosticId: "fixture", status: "warn", checks: [{ id: "mcp", status: "pass", code: "MCP_CATALOG_OK", message: "Schemas verified" }] })),
    updateInfo: vi.fn(async () => ({ status: { state: "done" } })) };
  return { info, client, panel: setupSystemPanel(() => client, () => busy) };
}
describe("Settings system and recovery controls", () => {
  it("shows private CLI versions, retained releases and sign-in instructions", async () => {
    const w = fixture(); await w.panel.refresh();
    expect(el("system-version").textContent).toContain("Previous: 1.0.0");
    expect(el("system-providers").textContent).toContain("private CLI 1.0.95");
    expect(el("system-providers").textContent).toContain("office-partner providers login codex");
    expect((el("system-rollback-version") as HTMLSelectElement).value).toBe("1.0.0");
  });
  it("prevents management actions while the current turn is active", async () => {
    const w = fixture(true); await w.panel.refresh(); el("system-repair").click(); await flush();
    expect(w.client.systemAction).not.toHaveBeenCalled(); expect(el("system-status").textContent).toContain("current response");
  });
  it("checks the actual host MCP catalog and displays stable diagnostic codes", async () => {
    const w = fixture(); el("system-check-mcp").click(); await flush();
    expect(w.client.systemDiagnostics).toHaveBeenCalledWith({ mcp: true });
    expect(el("system-report").textContent).toContain("MCP_CATALOG_OK");
  });
  it("reports failed provider activation and still refreshes pending sign-in state", async () => {
    const w = fixture(); w.client.systemAction.mockRejectedValueOnce(new Error("Private CLI needs sign-in"));
    await w.panel.refresh(); (el("system-providers").querySelectorAll("button")[1] as HTMLButtonElement).click(); await flush();
    expect(w.client.systemAction).toHaveBeenCalledWith("install", "codex");
    expect(el("system-status").textContent).toContain("needs sign-in"); expect(w.client.systemInfo).toHaveBeenCalledTimes(2);
  });
  it("requests the selected cached release and waits for recovery completion", async () => {
    vi.useFakeTimers(); vi.spyOn(window, "confirm").mockReturnValue(true);
    const w = fixture(); await w.panel.refresh(); el("system-rollback").click(); await flush();
    expect(w.client.systemAction).toHaveBeenCalledWith("rollback", undefined, "1.0.0");
    await vi.advanceTimersByTimeAsync(2000); expect(el("system-status").textContent).toBe("Recovery completed.");
  });
});
