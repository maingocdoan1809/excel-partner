import { describe, expect, it } from "vitest";
import {
  MAX_STORED_ARTIFACT,
  MAX_STORED_TEXT,
  buildContextPreface,
  clipSession,
  formatWhen,
  newSession,
  titleFrom,
  bindProviderSession,
  inferHost,
  type StoredMessage,
  type StoredSession,
} from "../addin/taskpane/ui/sessions.js";

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

describe("provider session binding", () => {
  it.each(["claude", "copilot", "codex", "opencode"])("keeps %s native titles and resumes the same history item", provider => {
    const first = bindProviderSession(chat("provisional", 10, [user("paint red")]),
      { provider, providerSessionId: "native-1", title: "Change Shape Color to Red" }, 0, 0, 10);
    expect(first.current).toMatchObject({ id: `${provider}:native-1`, title: "Change Shape Color to Red" });
    expect(first.previous?.messages).toEqual([]);
    first.current.messages.push(user("now blue"));
    const resumed = bindProviderSession(first.current, { provider, providerSessionId: "native-1" }, 1, 0, 20);
    expect(resumed.previous).toBeNull();
    expect(resumed.current).toMatchObject({ providerSessionId: "native-1", title: "Change Shape Color to Red" });
    const renamed = bindProviderSession(resumed.current,
      { provider, providerSessionId: "native-1", title: "Updated Native Title" }, 1, 0, 20);
    expect(renamed.current.title).toBe("Updated Native Title");
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
});

describe("clipSession", () => {
  it("cuts oversized text, tool fields and documents without changing the original or dropping messages", () => {
    const tool = { name: "read_range", args: "A1", result: "r".repeat(MAX_STORED_TEXT * 2), state: "ok" as const, error: null };
    const original = chat("big", 1, [
      user("x".repeat(MAX_STORED_TEXT * 2)),
      { kind: "tool_group", text: "", tools: [tool] },
      { kind: "artifact", text: "", artifact: { name: "report.md", content: "d".repeat(MAX_STORED_ARTIFACT * 2) } },
      ...Array.from({ length: 600 }, (_, i) => bot(`m${i}`)),
    ]);
    const clipped = clipSession(original);
    expect(clipped.messages).toHaveLength(603);
    expect(clipped.messages[0]!.text).toHaveLength(MAX_STORED_TEXT + 1);
    expect(clipped.messages[1]!.tools![0]!.result).toHaveLength(MAX_STORED_TEXT + 1);
    expect(clipped.messages[2]!.artifact!.content).toHaveLength(MAX_STORED_ARTIFACT + 1);
    expect(original.messages[0]!.text).toHaveLength(MAX_STORED_TEXT * 2);
    expect(clipSession(clipped)).toEqual(clipped);
  });
});

describe("history by app", () => {
  it("files chats saved before apps were recorded by what they used: Word or PowerPoint tools, otherwise Excel", () => {
    const withTool = (tool: string): StoredSession => ({ ...chat("t", 1, [user("q"), { kind: "tool", text: "", tool: { name: tool, args: "", state: "ok", error: null } }]) });
    expect(inferHost(chat("none", 1, [user("no tools at all")]))).toBe("excel");
    expect(inferHost(withTool("read_range"))).toBe("excel");
    expect(inferHost(withTool("edit_doc_text"))).toBe("word");
    expect(inferHost(withTool("read_slide_text"))).toBe("powerpoint");
    // Tools every app has say nothing about the app.
    expect(inferHost(withTool("get_selection"))).toBe("excel");
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
