// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { setupHistoryPanel } from "../frontend/taskpane/ui/history-panel.js";
import type { StoredSession } from "../protocol/src/history.js";

function failureFixture() {
 document.body.innerHTML = '<button id="clear"></button><div id="list"></div><div id="title"></div><div id="meta"></div><div id="transcript"></div><button id="continue"></button><textarea id="input"></textarea>';
 const session: StoredSession = { id: "saved", title: "Saved chat", provider: "copilot", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "keep this" }], changes: [] };
 const store = { list: () => [{ ...session, messageCount: 1, changeCount: 0 }], get: () => session, save() {}, latest: () => null,
  refresh: vi.fn(async () => {}), delete: vi.fn(async (_id: string) => {}), clear: vi.fn(async () => {}),
  loadPage: vi.fn(async (_id: string) => ({ session, revision: 1, generation: 0, totalMessages: 1, totalChanges: 0 })) };
 const element = (id: string) => document.getElementById(id)!;
 const resetChat = vi.fn(), abandonRequest = vi.fn(), showPanel = vi.fn();
 const panel = setupHistoryPanel({ store, getClient: () => ({ connected: false }) as never, getSession: () => session, getRequestId: () => null,
  persistNow() {}, abandonRequest, resetChat, refreshUsage: async () => {}, loadSession() {}, showPanel,
  providerNames: { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode" }, renderMessage: container => container,
  elements: { historyClear: element("clear") as HTMLButtonElement, historyList: element("list"), sessionTitle: element("title"), sessionMeta: element("meta"), sessionTranscript: element("transcript"), sessionContinue: element("continue") as HTMLButtonElement, input: element("input") as HTMLTextAreaElement } });
 return { panel, store, element, resetChat, abandonRequest, showPanel };
}

describe("history errors and acknowledgements", () => {
 it("H22 keeps the active chat until deletion is acknowledged", async () => {
  const w = failureFixture(); let acknowledge!: () => void;
  w.store.delete.mockImplementationOnce(() => new Promise<void>(resolve => { acknowledge = resolve; }));
  await w.panel.renderHistory(); w.element("list").querySelector<HTMLButtonElement>(".history-delete")!.click();
  expect(w.resetChat).not.toHaveBeenCalled(); expect(w.abandonRequest).not.toHaveBeenCalled();
  acknowledge(); await vi.waitFor(() => expect(w.resetChat).toHaveBeenCalledOnce());
 });
 it("H22 reports failed delete and clear without resetting the active chat", async () => {
  const w = failureFixture(); w.store.delete.mockRejectedValueOnce(new Error("offline"));
  await w.panel.renderHistory(); w.element("list").querySelector<HTMLButtonElement>(".history-delete")!.click();
  await vi.waitFor(() => expect(w.element("list").textContent).toContain("Could not delete chat: offline"));
  w.store.clear.mockRejectedValueOnce(new Error("disk full"));
  await w.panel.clearHistory(); await w.panel.clearHistory();
  expect(w.element("list").textContent).toContain("Could not clear history: disk full");
  expect(w.resetChat).not.toHaveBeenCalled(); expect(w.abandonRequest).not.toHaveBeenCalled();
 });
 it("H27 displays read failures and permits retry without claiming history is empty", async () => {
  const w = failureFixture(); w.store.refresh.mockRejectedValueOnce(new Error("offline"));
  await w.panel.renderHistory(); expect(w.element("list").textContent).toContain("History is unavailable: offline");
  expect(w.element("list").textContent).not.toContain("No saved chats");
  await w.panel.renderHistory(); w.store.loadPage.mockRejectedValueOnce(new Error("disconnected"));
  w.element("list").querySelector<HTMLButtonElement>(".history-item")!.click();
  await vi.waitFor(() => expect(w.element("list").textContent).toContain("Could not open history: disconnected"));
  expect(w.showPanel).not.toHaveBeenCalled();
  await w.panel.renderHistory(); w.element("list").querySelector<HTMLButtonElement>(".history-item")!.click();
  await vi.waitFor(() => expect(w.showPanel).toHaveBeenCalledWith("session"));
 });
});

it("H04 loads only the latest transcript page for browsing and the full chat only for Continue", async () => {
 document.body.innerHTML = '<button id="clear"></button><div id="list"></div><div id="title"></div><div id="meta"></div><div id="transcript"></div><button id="continue"></button><textarea id="input"></textarea>';
 const session: StoredSession = { id: "saved", title: "saved", provider: "copilot", startedAt: 1, updatedAt: 2, messages: Array.from({ length: 250 }, (_, i) => ({ kind: "user", text: String(i) })), changes: [] };
 const store = { list: () => [{ ...session, messageCount: 250, changeCount: 0 }], get: () => null, save() {}, delete() {}, clear() {}, latest: () => null,
 load: vi.fn(async () => session), loadPage: vi.fn(async (_id: string, offset = 0, limit = 100) => ({ session: { ...session, messages: session.messages.slice(offset, offset + limit) }, revision: 1, generation: 0, totalMessages: 250, totalChanges: 0 })) };
 const el = (id: string) => document.getElementById(id)!;
 const loadSession = vi.fn();
 const panel = setupHistoryPanel({ store, getClient: () => ({ connected: false }) as never, getSession: () => ({ ...session, id: "current" }), getRequestId: () => null, persistNow() {}, abandonRequest() {}, resetChat() {}, refreshUsage: async () => {}, loadSession, showPanel() {}, providerNames: { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode" }, renderMessage(container, message) { const node = document.createElement("p"); node.textContent = message.text; container.append(node); return node; }, elements: { historyClear: el("clear") as HTMLButtonElement, historyList: el("list"), sessionTitle: el("title"), sessionMeta: el("meta"), sessionTranscript: el("transcript"), sessionContinue: el("continue") as HTMLButtonElement, input: el("input") as HTMLTextAreaElement } });
 await panel.renderHistory(); (el("list").querySelector("button")!).click();
 await new Promise(resolve => setTimeout(resolve, 0));
 expect(el("transcript").querySelectorAll("p")).toHaveLength(100); expect(store.load).not.toHaveBeenCalled();
 expect(store.loadPage).toHaveBeenLastCalledWith("saved", 150, 100);
 await panel.continueViewedSession(); expect(store.load).toHaveBeenCalledWith("saved"); expect(loadSession).toHaveBeenCalledWith(session);
});
