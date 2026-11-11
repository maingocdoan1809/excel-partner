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
 const client = { historyStats: vi.fn(async () => ({ count: 1, bytes: 100, generation: 0 })) };
 const store = { pendingCount: 1, pendingExport: () => [{ id: "pending" }], refresh: vi.fn(async () => {}) };
 const history = setupHistoryStorage(() => client as never, store as never, () => {}, () => false);
 const click = (selector: string) => document.querySelector<HTMLButtonElement>(selector)!.click();
 return { client, nativeConfirm, click, history };
}
it("D02 offers one Delete all in the history section and no separate old-chat cleanup", async () => {
 const w = fixture(); await w.history.refresh();
 const section = document.getElementById("history-storage-section")!;
 expect(section.querySelectorAll(".settings-danger")).toHaveLength(1);
 expect(section.querySelector("#settings-clear-history")).not.toBeNull();
 expect(document.getElementById("history-cleanup")).toBeNull();
 expect(document.getElementById("history-storage-info")!.textContent).toContain("MB on disk");
 expect(w.client.historyStats).toHaveBeenCalled();
});
