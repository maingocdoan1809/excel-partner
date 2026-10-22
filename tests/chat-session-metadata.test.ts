import { describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import vm from "node:vm";
import ts from "typescript";
import { bindProviderSession, newSession, SessionStore } from "../addin/taskpane/ui/sessions.js";

// Exercise the task pane's actual event handler without starting Office or the DOM.
const main = fs.readFileSync(new URL("../addin/taskpane/main.ts", import.meta.url), "utf8").replace(/\r\n/g, "\n");
const start = main.indexOf("function handleEvent(");
const handler = ts.transpileModule(main.slice(start, main.indexOf("\n}\n", start) + 3), {
  compilerOptions: { target: ts.ScriptTarget.ES2020 },
}).outputText;

function setup() {
  const data = new Map<string, string>();
  const store = new SessionStore({
    getItem: key => data.get(key) ?? null,
    setItem: (key, value) => { data.set(key, value); },
    removeItem: key => { data.delete(key); },
  });
  const session = newSession("claude");
  session.messages.push({ kind: "user", text: "Keep this conversation" });
  const context = vm.createContext({
    session, requestSession: session, currentRequestId: "request", store, bindProviderSession,
    requestMessageIndex: 0, requestChangeIndex: 0, requestStartedAt: 1,
    usageRequestId: null, usageSummary: undefined,
    clearThinking: vi.fn(), showThinking: vi.fn(), renderChanges: vi.fn(), renderMessage: vi.fn(),
    els: { messages: { replaceChildren: vi.fn() } },
  });
  vm.runInContext(handler + "\nfunction persistNow() { store.save(session); }", context);
  const metadata = (title?: string, requestId = "request") => context.handleEvent({
    type: "session_info", requestId, provider: "claude", providerSessionId: "native", title,
  });
  return { context, store, metadata };
}

describe("task pane session metadata ownership", () => {
  it("ignores the first metadata if a new chat was opened before provider initialization", () => {
    const { context, store, metadata } = setup();
    const original = context.session;
    store.save(original);
    const next = newSession("claude", 2);
    context.session = next;
    metadata();
    store.save(context.session);
    expect(context.session).toBe(next);
    expect(store.get(original.id)?.messages).toEqual(original.messages);
    expect(store.get("claude:native")).toBeNull();
  });

  it("ignores late metadata after switching to a new or saved chat, preserving native history", () => {
    for (const useSavedChat of [false, true]) {
      const { context, store, metadata } = setup();
      metadata("Original title");
      const original = store.get("claude:native");
      const next = newSession("claude", 2);
      if (useSavedChat) next.messages.push({ kind: "user", text: "Another chat" });
      context.session = next;
      metadata("Late title");
      store.save(context.session);
      expect(context.session).toBe(next);
      expect(next.providerSessionId).toBeUndefined();
      expect(store.get("claude:native")).toEqual(original);
    }
  });

  it("accepts subsequent metadata after the active request binds its native ID", () => {
    const { context, store, metadata } = setup();
    metadata();
    expect(context.requestSession).toBe(context.session);
    metadata("Confirmed title");
    expect(store.get("claude:native")?.title).toBe("Confirmed title");
    metadata("Stale title", "older-request");
    expect(store.get("claude:native")?.title).toBe("Confirmed title");
  });
});
