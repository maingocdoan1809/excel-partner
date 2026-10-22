import { describe, expect, it } from "vitest";
import {
  LIMITS,
  SessionStore,
  buildContextPreface,
  formatWhen,
  newSession,
  titleFrom,
  bindProviderSession,
  type StoredMessage,
  type StoredSession,
} from "../addin/taskpane/ui/sessions.js";

class MemoryStorage {
  data = new Map<string, string>();
  quota = Infinity;
  getItem(key: string) {
    return this.data.get(key) ?? null;
  }
  setItem(key: string, value: string) {
    if (value.length > this.quota) throw new Error("QuotaExceededError");
    this.data.set(key, value);
  }
  removeItem(key: string) {
    this.data.delete(key);
  }
}

const user = (text: string): StoredMessage => ({ kind: "user", text });
const bot = (text: string): StoredMessage => ({ kind: "assistant", text });
const chat = (id: string, at: number, messages: StoredMessage[], changes: StoredSession["changes"] = []): StoredSession => ({
  id,
  title: "x",
  provider: "codex",
  startedAt: at,
  updatedAt: at,
  messages,
  changes,
});

describe("SessionStore", () => {
  it.each(["claude", "copilot", "codex", "opencode"])("keeps %s native titles and resumes the same history item", provider => {
    const store = new SessionStore(new MemoryStorage());
    const first = bindProviderSession(chat("provisional", 10, [user("paint red")]),
      { provider, providerSessionId: "native-1", title: "Change Shape Color to Red" }, 0, 0, 10);
    store.save(first.current);
    expect(store.list()[0]?.title).toBe("Change Shape Color to Red");
    expect(first.previous?.messages).toEqual([]);
    const restored = store.get(`${provider}:native-1`)!;
    restored.messages.push(user("now blue"));
    const resumed = bindProviderSession(restored, { provider, providerSessionId: "native-1" }, 1, 0, 20);
    expect(resumed.previous).toBeNull();
    store.save(resumed.current);
    expect(store.list()).toHaveLength(1);
    expect(store.get(`${provider}:native-1`)).toMatchObject({ providerSessionId: "native-1", title: "Change Shape Color to Red" });
    const renamed = bindProviderSession(resumed.current,
      { provider, providerSessionId: "native-1", title: "Updated Native Title" }, 1, 0, 20);
    store.save(renamed.current);
    expect(store.list()[0]?.title).toBe("Updated Native Title");
  });

  it("splits only the new request when the provider explicitly returns a different native session", () => {
    const original = { ...chat("copilot:old", 10, [user("old task"), bot("done"), user("new task")]),
      provider: "copilot", providerSessionId: "old", providerTitle: "Old Native Title" };
    const result = bindProviderSession(original, { provider: "copilot", providerSessionId: "new" }, 2, 0, 30);
    expect(result.previous?.messages).toHaveLength(2);
    expect(result.previous?.providerTitle).toBe("Old Native Title");
    expect(result.current.id).toBe("copilot:new");
    expect(result.current.title).toBe("new task");
    expect(result.current.messages).toEqual([user("new task")]);
    expect(result.current.startedAt).toBe(30);
  });
  it("saves, lists newest first and loads a session back", () => {
    const store = new SessionStore(new MemoryStorage());
    store.save(chat("a", 1000, [user("first question"), bot("answer")]));
    store.save(chat("b", 2000, [user("second"), bot("ok"), user("more")], [{ at: 1, tool: "write_range", ok: true, summary: "Wrote", target: null }]));
    const list = store.list();
    expect(list.map((s) => s.id)).toEqual(["b", "a"]);
    expect(list[0]).toMatchObject({ title: "second", messageCount: 3, changeCount: 1, provider: "codex" });
    expect(store.get("a")!.messages).toHaveLength(2);
    expect(store.get("zzz")).toBeNull();
  });

  it("does not keep or list chats without a user message", () => {
    const store = new SessionStore(new MemoryStorage());
    store.save(chat("empty", 1, []));
    store.save(chat("only-bot", 2, [bot("hi")]));
    expect(store.list()).toEqual([]);
  });

  it("updates in place, deletes one and clears all", () => {
    const store = new SessionStore(new MemoryStorage());
    store.save(chat("a", 1, [user("q")]));
    store.save({ ...chat("a", 5, [user("q"), bot("a")]) });
    expect(store.list()).toHaveLength(1);
    expect(store.get("a")!.updatedAt).toBe(5);
    store.save(chat("b", 6, [user("q2")]));
    store.delete("a");
    expect(store.list().map((s) => s.id)).toEqual(["b"]);
    store.clear();
    expect(store.list()).toEqual([]);
  });

  it("returns the latest chat only while it is recent enough", () => {
    const store = new SessionStore(new MemoryStorage());
    store.save(chat("old", 1_000, [user("q")]));
    expect(store.latest(60_000, 30_000)?.id).toBe("old");
    expect(store.latest(60_000, 500_000)).toBeNull();
  });

  it("restores tool call groups without old update notices", () => {
    const store = new SessionStore(new MemoryStorage());
    const group: StoredMessage = { kind: "tool_group", text: "", tools: [{ name: "read_range", args: "A1", state: "ok", error: null }] };
    store.save(chat("tools", 100, [user("read"), group, bot("done"), { kind: "note", text: "Updating Excel Partner…" }]));
    expect(store.get("tools")!.messages).toEqual([user("read"), group, bot("done")]);
  });

  it("survives corrupt storage", () => {
    const storage = new MemoryStorage();
    storage.setItem("excelPartner.sessions.v1", "{not json");
    const store = new SessionStore(storage);
    expect(store.list()).toEqual([]);
    store.save(chat("a", 1, [user("q")]));
    expect(store.list()).toHaveLength(1);
  });

  it("caps sessions, messages and message length", () => {
    const store = new SessionStore(new MemoryStorage());
    for (let i = 0; i < LIMITS.sessions + 5; i++) store.save(chat(`s${i}`, i + 1, [user(`q${i}`)]));
    expect(store.list()).toHaveLength(LIMITS.sessions);
    expect(store.list()[0]!.id).toBe(`s${LIMITS.sessions + 4}`);

    const many = Array.from({ length: LIMITS.messagesPerSession + 50 }, (_, i) => (i % 2 === 0 ? user(`m${i}`) : bot(`m${i}`)));
    store.save(chat("big", 10_000, many));
    expect(store.get("big")!.messages).toHaveLength(LIMITS.messagesPerSession);
    expect(store.get("big")!.title).toBe("m0");
    store.save(chat("long", 10_001, [user("x".repeat(LIMITS.textChars * 2))]));
    expect(store.get("long")!.messages[0]!.text.length).toBeLessThanOrEqual(LIMITS.textChars + 1);
  });

  it("falls back to the newest chat when the quota is exceeded", () => {
    const storage = new MemoryStorage();
    const store = new SessionStore(storage);
    store.save(chat("a", 1, [user("old chat")]));
    storage.quota = 250; // too small for two sessions
    store.save(chat("b", 2, [user("new chat")]));
    expect(store.list().map((s) => s.id)).toEqual(["b"]);
  });
});

describe("titles and context", () => {
  it("titles a chat with its first user message, one line, truncated", () => {
    expect(titleFrom([bot("hi"), user("  vẽ flow\nlogin  ")])).toBe("vẽ flow login");
    expect(titleFrom([user("x".repeat(100))])).toHaveLength(68);
    expect(titleFrom([])).toBe("New chat");
    expect(newSession("codex").title).toBe("New chat");
  });

  it("replays the conversation and the changes made, newest part first when over budget", () => {
    const preface = buildContextPreface({
      messages: [user("make a flow"), bot("done: 5 shapes"), { kind: "tool", text: "", tool: { name: "add_shape", args: "{}", state: "ok", error: null } }, user("now colour it")],
      changes: [
        { at: 1, tool: "add_shape", ok: true, summary: 'Added Rectangle "a"', target: null },
        { at: 2, tool: "update_shape", ok: false, summary: "update_shape failed", target: null },
      ],
    });
    expect(preface).toContain("<previous_conversation");
    expect(preface).toContain("User: make a flow\nAssistant: done: 5 shapes\nUser: now colour it");
    expect(preface).toContain('- Added Rectangle "a"');
    expect(preface).not.toContain("update_shape failed");
    expect(preface).not.toContain("add_shape");

    const long = Array.from({ length: 60 }, (_, i) => (i % 2 ? bot("a".repeat(500)) : user(`question ${i}`)));
    const trimmed = buildContextPreface({ messages: long, changes: [] }, 3_000);
    expect(trimmed).toContain("(earlier messages omitted)");
    expect(trimmed).toContain("question 58");
    expect(trimmed).not.toContain("question 0\n");
    expect(trimmed.length).toBeLessThan(4_500);
  });
});

describe("formatWhen", () => {
  const now = new Date(2026, 9, 14, 12, 0, 0).getTime();
  it("uses relative wording for recent times and dates for older ones", () => {
    expect(formatWhen(now - 10_000, now)).toBe("Just now");
    expect(formatWhen(now - 5 * 60_000, now)).toBe("5 min ago");
    expect(formatWhen(new Date(2026, 9, 14, 8, 5).getTime(), now)).toBe("Today 08:05");
    expect(formatWhen(new Date(2026, 9, 13, 14, 5).getTime(), now)).toBe("Yesterday 14:05");
    expect(formatWhen(new Date(2026, 9, 1, 9, 30).getTime(), now)).toBe("1 Oct 09:30");
  });
});
