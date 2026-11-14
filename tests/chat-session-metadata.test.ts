// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from "vitest";
import { handleEvent } from "../frontend/taskpane/chat/events.js";
import { app } from "../frontend/taskpane/chat/state.js";
import { newSession, type StoredSession } from "../frontend/taskpane/ui/sessions.js";

// Exercise the task pane's real event handler. Only what draws on the page or talks to the bridge is replaced.
const fakes = vi.hoisted(() => {
  const saved = new Map<string, unknown>();
  return {
    saved,
    // Stands in for DurableHistory: a chat is kept once it has a user message, and a saved chat is returned as a copy.
    store: {
      save: (chat: { id: string; messages: { kind: string }[] }) => { if (chat.messages.some(m => m.kind === "user")) saved.set(chat.id, structuredClone(chat)); },
      get: (id: string) => saved.get(id) ?? null,
      flush: async () => {},
    },
    usagePanel: { rebind: vi.fn() },
    thinking: { clear: vi.fn(), show: vi.fn(), setLabel: vi.fn() },
  };
});
vi.mock("../frontend/taskpane/chat/services.js", () => ({ store: fakes.store, usagePanel: fakes.usagePanel, thinking: fakes.thinking, renderChanges: vi.fn() }));
vi.mock("../frontend/taskpane/chat/messages.js", () => ({ renderMessage: vi.fn() }));
vi.mock("../frontend/taskpane/chat/turn.js", () => ({}));
vi.mock("../frontend/taskpane/ui/elements.js", () => ({ els: { messages: { replaceChildren: vi.fn() } } }));
vi.mock("../frontend/taskpane/chat/chats.js", async () => {
  const { app } = await import("../frontend/taskpane/chat/state.js");
  return { persistNow: () => fakes.store.save(app.session), schedulePersist: () => {} };
});

function setup() {
  fakes.saved.clear();
  fakes.usagePanel.rebind.mockClear();
  const session = newSession("claude");
  session.messages.push({ kind: "user", text: "Keep this conversation" });
  Object.assign(app, { session, requestSession: session, currentRequestId: "request", requestMessageIndex: 0, requestChangeIndex: 0, requestStartedAt: 1 });
  const metadata = (title?: string, requestId = "request") => handleEvent({
    type: "session_info", requestId, provider: "claude", providerSessionId: "native", title,
  });
  const saved = (id: string) => fakes.store.get(id) as StoredSession | null;
  return { store: { ...fakes.store, get: saved }, usagePanel: fakes.usagePanel, metadata };
}

describe("task pane session metadata ownership", () => {
  beforeEach(() => { app.session = newSession("claude"); });

  it("ignores the first metadata if a new chat was opened before provider initialization", () => {
    const { store, usagePanel, metadata } = setup();
    const original = app.session;
    store.save(original);
    const next = newSession("claude", 2);
    app.session = next;
    metadata();
    store.save(app.session);
    expect(app.session).toBe(next);
    expect(store.get(original.id)?.messages).toEqual(original.messages);
    expect(store.get("claude:native")).toBeNull();
    expect(usagePanel.rebind).not.toHaveBeenCalled();
  });

  it("ignores late metadata after switching to a new or saved chat, preserving native history", () => {
    for (const useSavedChat of [false, true]) {
      const { store, usagePanel, metadata } = setup();
      metadata("Original title");
      const original = store.get("claude:native");
      const next = newSession("claude", 2);
      if (useSavedChat) next.messages.push({ kind: "user", text: "Another chat" });
      app.session = next;
      metadata("Late title");
      store.save(app.session);
      expect(app.session).toBe(next);
      expect(next.providerSessionId).toBeUndefined();
      expect(store.get("claude:native")).toEqual(original);
      expect(usagePanel.rebind).toHaveBeenCalledTimes(1);
    }
  });

  it("accepts subsequent metadata after the active request binds its native ID", () => {
    const { store, usagePanel, metadata } = setup();
    metadata();
    expect(app.requestSession).toBe(app.session);
    metadata("Confirmed title");
    expect(store.get("claude:native")?.title).toBe("Confirmed title");
    metadata("Stale title", "older-request");
    expect(store.get("claude:native")?.title).toBe("Confirmed title");
    expect(usagePanel.rebind).toHaveBeenCalledTimes(1);
    expect(usagePanel.rebind).toHaveBeenCalledWith("request");
  });
});
