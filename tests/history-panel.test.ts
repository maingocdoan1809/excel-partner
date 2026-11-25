// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { setupHistoryPanel } from "../frontend/taskpane/ui/history-panel.js";
import type { StoredSession } from "../protocol/src/history.js";

function failureFixture() {
 document.body.innerHTML = '<button id="clear"></button><button id="import"></button><div id="import-list" hidden></div><div id="list"></div><div id="title"></div><div id="meta"></div><div id="transcript"></div><button id="continue"></button><textarea id="input"></textarea>';
 const session: StoredSession = { id: "saved", title: "Saved chat", provider: "copilot", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "keep this" }], changes: [] };
 const store = { list: () => [{ ...session, messageCount: 1, changeCount: 0 }], get: () => session, save() {}, latest: () => null, latestForDocument: () => null,
  refresh: vi.fn(async () => {}), delete: vi.fn(async (_id: string) => {}), clear: vi.fn(async () => {}),
  loadPage: vi.fn(async (_id: string) => ({ session, revision: 1, generation: 0, totalMessages: 1, totalChanges: 0 })) };
 const element = (id: string) => document.getElementById(id)!;
 const resetChat = vi.fn(), abandonRequest = vi.fn(), showPanel = vi.fn();
 const panel = setupHistoryPanel({ store, getClient: () => ({ connected: false }) as never, getSession: () => session, getRequestId: () => null, getProvider: () => "copilot" as const, canImport: () => true,
  persistNow() {}, abandonRequest, resetChat, refreshUsage: async () => {}, loadSession() {}, showPanel,
  providerNames: { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode", antigravity: "Antigravity" }, renderMessage: container => container,
  elements: { historyClear: element("clear") as HTMLButtonElement, historyImport: element("import") as HTMLButtonElement, historyImportList: element("import-list"), historyList: element("list"), sessionTitle: element("title"), sessionMeta: element("meta"), sessionTranscript: element("transcript"), sessionContinue: element("continue") as HTMLButtonElement, input: element("input") as HTMLTextAreaElement } });
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
 document.body.innerHTML = '<button id="clear"></button><button id="import"></button><div id="import-list" hidden></div><div id="list"></div><div id="title"></div><div id="meta"></div><div id="transcript"></div><button id="continue"></button><textarea id="input"></textarea>';
 const session: StoredSession = { id: "saved", title: "saved", provider: "copilot", startedAt: 1, updatedAt: 2, messages: Array.from({ length: 250 }, (_, i) => ({ kind: "user", text: String(i) })), changes: [] };
 const store = { list: () => [{ ...session, messageCount: 250, changeCount: 0 }], get: () => null, save() {}, delete() {}, clear() {}, latest: () => null, latestForDocument: () => null,
 load: vi.fn(async () => session), loadPage: vi.fn(async (_id: string, offset = 0, limit = 100) => ({ session: { ...session, messages: session.messages.slice(offset, offset + limit) }, revision: 1, generation: 0, totalMessages: 250, totalChanges: 0 })) };
 const el = (id: string) => document.getElementById(id)!;
 const loadSession = vi.fn();
 const panel = setupHistoryPanel({ store, getClient: () => ({ connected: false }) as never, getSession: () => ({ ...session, id: "current" }), getRequestId: () => null, getProvider: () => "copilot" as const, canImport: () => true, persistNow() {}, abandonRequest() {}, resetChat() {}, refreshUsage: async () => {}, loadSession, showPanel() {}, providerNames: { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode", antigravity: "Antigravity" }, renderMessage(container, message) { const node = document.createElement("p"); node.textContent = message.text; container.append(node); return node; }, elements: { historyClear: el("clear") as HTMLButtonElement, historyImport: el("import") as HTMLButtonElement, historyImportList: el("import-list"), historyList: el("list"), sessionTitle: el("title"), sessionMeta: el("meta"), sessionTranscript: el("transcript"), sessionContinue: el("continue") as HTMLButtonElement, input: el("input") as HTMLTextAreaElement } });
 await panel.renderHistory(); (el("list").querySelector("button")!).click();
 await new Promise(resolve => setTimeout(resolve, 0));
 expect(el("transcript").querySelectorAll("p")).toHaveLength(100); expect(store.load).not.toHaveBeenCalled();
 expect(store.loadPage).toHaveBeenLastCalledWith("saved", 150, 100);
 await panel.continueViewedSession(); expect(store.load).toHaveBeenCalledWith("saved"); expect(loadSession).toHaveBeenCalledWith(session);
});

describe("importing conversations from the provider's own tools", () => {
 const item = { id: "vscode-chat-1", title: "Closures", origin: "VS Code", cwd: "/work/app", updatedAt: 5, messageCount: 2 };
 function importFixture(sessions: unknown[]) {
  const w = failureFixture();
  const client = { connected: true, externalSessions: vi.fn(async (_provider: string) => ({ sessions })), importExternalSession: vi.fn(async (_provider: string, _id: string) => ({ id: "copilot:native-1", existing: false })) };
  const panel = setupHistoryPanel({ store: w.store, getClient: () => client as never, getSession: () => ({ id: "current", title: "", provider: "copilot", startedAt: 1, updatedAt: 1, messages: [], changes: [] }), getRequestId: () => null,
   getProvider: () => "copilot" as const, canImport: () => true, persistNow() {}, abandonRequest() {}, resetChat() {}, refreshUsage: async () => {}, loadSession() {}, showPanel: w.showPanel,
   providerNames: { claude: "Claude", copilot: "Copilot", codex: "Codex", opencode: "OpenCode", antigravity: "Antigravity" }, renderMessage: container => container,
   elements: { historyClear: w.element("clear") as HTMLButtonElement, historyImport: w.element("import") as HTMLButtonElement, historyImportList: w.element("import-list"), historyList: w.element("list"), sessionTitle: w.element("title"), sessionMeta: w.element("meta"), sessionTranscript: w.element("transcript"), sessionContinue: w.element("continue") as HTMLButtonElement, input: w.element("input") as HTMLTextAreaElement } });
  return { ...w, client, panel };
 }
 it("lists them in place of the saved chats, imports the one picked and opens it to continue", async () => {
  const w = importFixture([item]); await w.panel.renderHistory();
  expect(w.element("import").hidden).toBe(false);
  await w.panel.toggleImport();
  expect(w.element("list").hidden).toBe(true); expect(w.element("import-list").hidden).toBe(false);
  expect(w.client.externalSessions).toHaveBeenCalledWith("copilot");
  expect(w.element("import-list").textContent).toContain("Closures"); expect(w.element("import-list").textContent).toContain("VS Code");
  w.element("import-list").querySelector<HTMLButtonElement>(".history-item")!.click();
  await vi.waitFor(() => expect(w.showPanel).toHaveBeenCalledWith("session"));
  expect(w.client.importExternalSession).toHaveBeenCalledWith("copilot", "vscode-chat-1");
  expect(w.store.loadPage).toHaveBeenCalledWith("copilot:native-1", 0, 100);
 });
 it("says so when there is nothing to import and shows a failed import without losing the list", async () => {
  const empty = importFixture([]); await empty.panel.renderHistory(); await empty.panel.toggleImport();
  expect(empty.element("import-list").textContent).toContain("No Copilot conversations");
  const w = importFixture([item]); await w.panel.renderHistory(); await w.panel.toggleImport();
  w.client.importExternalSession.mockRejectedValueOnce(new Error("not found"));
  w.element("import-list").querySelector<HTMLButtonElement>(".history-item")!.click();
  await vi.waitFor(() => expect(w.element("import-list").textContent).toContain("Could not import: not found"));
  expect(w.element("import-list").querySelector<HTMLButtonElement>(".history-item")!.disabled).toBe(false);
  await w.panel.toggleImport();
  expect(w.element("list").hidden).toBe(false); expect(w.element("import-list").hidden).toBe(true);
 });
});
