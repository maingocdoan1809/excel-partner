// @vitest-environment jsdom
import fs from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { setupSystemPanel, type SystemInfo } from "../frontend/taskpane/ui/system-panel.js";

const el = (id: string) => document.getElementById(id)!;
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
afterEach(() => { vi.useRealTimers(); vi.restoreAllMocks(); });
function fixture(busy = false) {
  document.body.innerHTML = fs.readFileSync("frontend/taskpane/index.html", "utf8").match(/<body[^>]*>([\s\S]*)<\/body>/)![1]!;
  el("panel-settings").hidden = false;
  const info: SystemInfo = { version: "1.1.0",
    providers: [{ provider: "copilot", selected: true, disabled: false, managed: true, version: "1.0.95", failure: null, available: true },
      { provider: "codex", selected: false, disabled: false, managed: false, version: null, failure: "NATIVE_PREFLIGHT_POLICY_DENIED", available: null }] };
  const client = { systemInfo: vi.fn(async () => info), systemAction: vi.fn(async (_action: string, _provider?: string, _version?: string) => {}),
    systemDiagnostics: vi.fn(async () => ({ diagnosticId: "fixture", status: "warn", checks: [{ id: "mcp", status: "pass", code: "MCP_CATALOG_OK", message: "Schemas verified" }] })),
    updateInfo: vi.fn(async () => ({ status: { state: "done" } })) };
  return { info, client, panel: setupSystemPanel(() => client, () => busy) };
}
describe("Settings system and recovery controls", () => {
  it("shows the installed version, private CLI versions and sign-in instructions", async () => {
    const w = fixture(); await w.panel.refresh();
    expect(el("system-version").textContent).toContain("Installed: 1.1.0");
    expect(el("system-version").textContent).not.toContain("Previous");
    expect(el("system-providers").textContent).toContain("private CLI 1.0.95");
    expect(el("system-providers").textContent).toContain("last installation did not finish");
  });
  it("prevents management actions while the current turn is active", async () => {
    const w = fixture(true); await w.panel.refresh(); el("system-repair").click(); await flush();
    expect(w.client.systemAction).not.toHaveBeenCalled(); expect(el("system-status").textContent).toContain("current response");
  });
  it("checks the actual host MCP catalog and displays stable diagnostic codes", async () => {
    const w = fixture(); el("system-check").click(); await flush();
    expect(w.client.systemDiagnostics).toHaveBeenCalledWith({ mcp: true });
    expect(el("system-report").textContent).toContain("MCP_CATALOG_OK");
  });
  it("offers export only after a check, and has no rollback control because old releases are not kept", async () => {
    const w = fixture(); await w.panel.refresh();
    expect(document.getElementById("system-rollback")).toBeNull(); expect(document.getElementById("system-rollback-version")).toBeNull();
    expect((el("system-export") as HTMLButtonElement).hidden).toBe(true);
    el("system-check").click(); await flush();
    expect((el("system-export") as HTMLButtonElement).hidden).toBe(false);
  });
  it("uses one button to start and stop debug logging, following the state the bridge reports", async () => {
    const w = fixture(); await w.panel.refresh();
    expect(el("system-debug").textContent).toBe("Debug for 15 min");
    el("system-debug").click(); await flush();
    expect(w.client.systemAction).toHaveBeenLastCalledWith("debug-on"); expect(el("system-debug").textContent).toBe("Stop debug");
    el("system-debug").click(); await flush();
    expect(w.client.systemAction).toHaveBeenLastCalledWith("debug-off"); expect(el("system-debug").textContent).toBe("Debug for 15 min");
    w.info.debug = true; await w.panel.refresh(); expect(el("system-debug").textContent).toBe("Stop debug");
  });
  it("D01 highlights failed checks in the summary and identifies their service", async () => {
    const w = fixture();
    w.client.systemDiagnostics.mockResolvedValueOnce({ diagnosticId: "fixture", status: "fail", checks: [{ id: "gateway.health", status: "fail", code: "SERVICE_NOT_READY", message: "Gateway on port 8766 is responding but still initializing." }] });
    el("system-check").click(); await flush();
    expect(el("system-status").textContent).toContain("1 check failed: gateway.health");
    expect(el("system-status").dataset.state).toBe("fail");
    expect(el("system-report").textContent).toContain("[fail] gateway.health · SERVICE_NOT_READY");
  });
  it("reports failed provider activation and still refreshes pending sign-in state", async () => {
    const w = fixture(); w.client.systemAction.mockRejectedValueOnce(new Error("Private CLI needs sign-in"));
    await w.panel.refresh(); (el("system-providers").querySelectorAll("button")[1] as HTMLButtonElement).click(); await flush();
    expect(w.client.systemAction).toHaveBeenCalledWith("install", "codex");
    expect(el("system-status").textContent).toContain("needs sign-in"); expect(w.client.systemInfo).toHaveBeenCalledTimes(2);
  });
  it("asks before repairing, repairs without native dialogs and waits for the installation to come back", async () => {
    vi.useFakeTimers(); const nativeConfirm = vi.spyOn(window, "confirm").mockImplementation(() => { throw Error("Unsupported"); });
    const w = fixture(); await w.panel.refresh(); el("system-repair").click(); await flush();
    expect(w.client.systemAction).not.toHaveBeenCalled();
    document.querySelector<HTMLButtonElement>("[data-confirm-accept]")!.click(); await flush();
    expect(w.client.systemAction).toHaveBeenCalledWith("repair");
    await vi.advanceTimersByTimeAsync(2000); expect(el("system-status").textContent).toBe("Recovery completed.");
    expect(nativeConfirm).not.toHaveBeenCalled();
  });
  it("D02 does nothing when the person cancels the repair", async () => {
    const w = fixture(); await w.panel.refresh(); el("system-repair").click(); await flush();
    document.querySelector<HTMLButtonElement>("[data-confirm-cancel]")!.click(); await flush();
    expect(w.client.systemAction).not.toHaveBeenCalled();
  });
});
