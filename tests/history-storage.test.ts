// @vitest-environment jsdom
import fs from "node:fs";
import { afterEach, expect, it, vi } from "vitest";
import { setupHistoryStorage } from "../addin/taskpane/ui/history-storage.js";
const tick = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
afterEach(() => { vi.restoreAllMocks(); });
function fixture() {
 document.body.innerHTML = fs.readFileSync("addin/taskpane/index.html", "utf8").match(/<body[^>]*>([\s\S]*)<\/body>/)![1]!;
 document.getElementById("panel-settings")!.hidden = false;
 const nativeConfirm = vi.spyOn(window, "confirm").mockImplementation(() => { throw Error("Function window.confirm is not supported."); });
 const client = { historyStats: vi.fn(async () => ({ count: 1, bytes: 100, generation: 0 })), historyCleanup: vi.fn(async (_before: number, _ids: string[]) => ({ removed: 0 })) };
 const store = { pendingCount: 1, pendingExport: () => [{ id: "pending" }], refresh: vi.fn(async () => {}) };
 setupHistoryStorage(() => client as never, store as never, () => "active", () => {}, () => false);
 const click = (selector: string) => document.querySelector<HTMLButtonElement>(selector)!.click();
 return { client, nativeConfirm, click };
}
it("D02 deletes old chats only after inline confirmation when Office blocks native confirm", async () => {
 const w = fixture(); w.click("#history-cleanup"); await tick();
 expect(w.client.historyCleanup).not.toHaveBeenCalled(); expect(document.querySelector(".inline-confirmation")).not.toBeNull();
 w.click("[data-confirm-cancel]"); await tick(); expect(w.client.historyCleanup).not.toHaveBeenCalled();
 w.click("#history-cleanup"); await tick(); w.click("[data-confirm-accept]"); await tick();
 expect(w.client.historyCleanup).toHaveBeenCalledWith(expect.any(Number), ["active", "pending"]);
 expect(w.nativeConfirm).not.toHaveBeenCalled(); expect(document.getElementById("history-storage-status")!.textContent).toContain("Deleted 0");
});
it("D02 cancels deletion when its age range changes or the panel closes", async () => {
 const w = fixture(); w.click("#history-cleanup"); await tick();
 (document.getElementById("history-cleanup-days") as HTMLInputElement).value = "30";
 w.click("[data-confirm-accept]"); await tick(); expect(w.client.historyCleanup).not.toHaveBeenCalled();
 w.click("#history-cleanup"); await tick(); document.getElementById("panel-settings")!.hidden = true; await tick();
 expect(w.client.historyCleanup).not.toHaveBeenCalled(); expect(document.querySelector(".inline-confirmation")).toBeNull();
});
