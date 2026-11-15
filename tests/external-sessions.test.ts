import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { parse } from "yaml";
import { HistoryStore } from "../bridge/src/history-store.js";
import type { StoredSession } from "../protocol/src/history.js";
import type { SessionHomes } from "../providers/src/external-sessions.js";
import { importClaudeSession, listClaudeSessions } from "../providers/src/claude/external.js";
import { importCodexSession, listCodexSessions } from "../providers/src/codex/external.js";
import { importCopilotSession, listCopilotSessions } from "../providers/src/copilot/external.js";
import { ensureWorkspaceDir } from "../providers/src/workspace.js";

const dirs: string[] = [];
let savedHome: string | undefined;
beforeEach(() => { savedHome = process.env.EXCEL_AI_HOME; process.env.EXCEL_AI_HOME = temp("scratch-home-"); });
afterEach(() => {
  if (savedHome === undefined) delete process.env.EXCEL_AI_HOME; else process.env.EXCEL_AI_HOME = savedHome;
  for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true });
});

function temp(prefix: string): string { const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix)); dirs.push(dir); return dir; }
function write(file: string, content: string): void { fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, content); }
const jsonl = (lines: unknown[]) => lines.map((line) => JSON.stringify(line)).join("\n") + "\n";
/** A person's own home and the private state folder of the add-in's CLI, which is where an import has to make the conversation resumable. */
function homes(extra: Record<string, (root: string) => string>, targetExtra: Record<string, (root: string) => string>): { homes: SessionHomes; person: string; product: string } {
  const person = temp("person-"), product = temp("product-");
  const env = (root: string, vars: Record<string, (root: string) => string>) => ({ HOME: root, USERPROFILE: root, APPDATA: root, XDG_CONFIG_HOME: root,
    ...Object.fromEntries(Object.entries(vars).map(([name, value]) => [name, value(root)])) });
  return { homes: { source: env(person, extra), target: env(product, targetExtra) }, person, product };
}
const encode = (cwd: string) => cwd.replace(/[^a-zA-Z0-9]/g, "-");

describe("Claude Code conversations from the terminal or VS Code", () => {
  const ID = "6c1f1a2e-0000-4000-8000-000000000001";
  const transcript = [
    { type: "user", cwd: "/work/app", entrypoint: "claude-vscode", timestamp: "2026-01-01T00:00:00Z", message: { role: "user", content: "Fix the bug\n<system-reminder>ignore me</system-reminder>" } },
    { type: "user", isMeta: true, message: { role: "user", content: "meta text" } },
    { type: "assistant", timestamp: "2026-01-01T00:00:05Z", message: { role: "assistant", content: [{ type: "thinking", thinking: "" }, { type: "text", text: "Looking." }, { type: "tool_use", id: "t1", name: "Read", input: { file_path: "a.ts" } }] } },
    { type: "user", message: { role: "user", content: [{ type: "tool_result", tool_use_id: "t1", content: [{ type: "text", text: "file text" }] }] } },
    { type: "assistant", isSidechain: true, message: { role: "assistant", content: [{ type: "text", text: "sub-agent chatter" }] } },
    { type: "assistant", timestamp: "2026-01-01T00:00:09Z", message: { role: "assistant", content: [{ type: "text", text: "Done." }] } },
    { type: "ai-title", aiTitle: "Fix the login bug" },
  ];
  const setup = () => {
    const h = homes({ CLAUDE_CONFIG_DIR: (root) => path.join(root, "claude") }, { CLAUDE_CONFIG_DIR: (root) => path.join(root, "claude-private") });
    write(path.join(h.person, "claude", "projects", encode("/work/app"), `${ID}.jsonl`), jsonl(transcript));
    return h;
  };

  it("lists what was held outside the add-in, named and placed, and leaves out the add-in's own", async () => {
    const h = setup();
    write(path.join(h.person, "claude", "projects", encode(ensureWorkspaceDir()), "6c1f1a2e-0000-4000-8000-000000000009.jsonl"), jsonl(transcript));
    const list = await listClaudeSessions(h.homes.source);
    expect(list).toEqual([expect.objectContaining({ id: ID, title: "Fix the login bug", origin: "VS Code", cwd: "/work/app" })]);
  });

  it("reads the words, the tools and their results, and copies the transcript to where the add-in's Claude resumes from", async () => {
    const h = setup();
    const imported = await importClaudeSession(h.homes, ID);
    expect(imported.providerSessionId).toBe(ID);
    expect(imported.title).toBe("Fix the login bug");
    expect(imported.messages).toEqual([
      { kind: "user", text: "Fix the bug" },
      { kind: "assistant", text: "Looking." },
      { kind: "tool_group", text: "", tools: [{ name: "Read", args: '{"file_path":"a.ts"}', result: "file text", state: "ok", error: null }] },
      { kind: "assistant", text: "Done." },
    ]);
    const copy = path.join(h.product, "claude-private", "projects", encode(ensureWorkspaceDir()), `${ID}.jsonl`);
    expect(fs.readFileSync(copy, "utf8")).toBe(fs.readFileSync(path.join(h.person, "claude", "projects", encode("/work/app"), `${ID}.jsonl`), "utf8"));
    await expect(importClaudeSession(h.homes, ID)).resolves.toMatchObject({ providerSessionId: ID });
  });

  it("refuses an id that is not a conversation id or that does not exist", async () => {
    const h = setup();
    await expect(importClaudeSession(h.homes, "../escape")).rejects.toThrow("Invalid");
    await expect(importClaudeSession(h.homes, "6c1f1a2e-0000-4000-8000-0000000000ff")).rejects.toThrow("not found");
  });
});

describe("Codex conversations from the terminal or VS Code", () => {
  const ID = "0199a213-aaaa-4bbb-8ccc-000000000001";
  const rollout = [
    { timestamp: "2026-01-01T00:00:00Z", type: "session_meta", payload: { id: ID, cwd: "/work/app", source: "vscode" } },
    { type: "response_item", payload: { type: "message", role: "user", content: [{ type: "input_text", text: "<environment_context>cwd</environment_context>" }] } },
    { timestamp: "2026-01-01T00:00:01Z", type: "event_msg", payload: { type: "user_message", message: "list the files" } },
    { type: "response_item", payload: { type: "function_call", name: "shell", arguments: '{"command":["ls"]}', call_id: "c1" } },
    { type: "response_item", payload: { type: "function_call_output", call_id: "c1", output: '{"output":"a.txt\\n","metadata":{}}' } },
    { timestamp: "2026-01-01T00:00:03Z", type: "event_msg", payload: { type: "agent_message", message: "Found a.txt" } },
  ];
  const setup = () => {
    const h = homes({ CODEX_HOME: (root) => path.join(root, "codex") }, { CODEX_HOME: (root) => path.join(root, "codex-private") });
    write(path.join(h.person, "codex", "sessions", "2026", "01", "01", `rollout-2026-01-01T00-00-00-${ID}.jsonl`), jsonl(rollout));
    write(path.join(h.person, "codex", "session_index.jsonl"), jsonl([{ id: ID, thread_name: "List files" }]));
    return h;
  };

  it("lists conversations with the name the CLI gave them and where they were held", async () => {
    const h = setup();
    expect(await listCodexSessions(h.homes.source)).toEqual([expect.objectContaining({ id: ID, title: "List files", origin: "VS Code", cwd: "/work/app" })]);
  });

  it("skips what Codex injects as the first message, joins tool calls with their output, and copies the rollout into a private state folder", async () => {
    const h = setup();
    const imported = await importCodexSession(h.homes, ID);
    expect(imported.messages).toEqual([
      { kind: "user", text: "list the files" },
      { kind: "tool_group", text: "", tools: [{ name: "shell", args: '{"command":["ls"]}', result: "a.txt\n", state: "ok", error: null }] },
      { kind: "assistant", text: "Found a.txt" },
    ]);
    expect(fs.existsSync(path.join(h.product, "codex-private", "sessions", "2026", "01", "01", `rollout-2026-01-01T00-00-00-${ID}.jsonl`))).toBe(true);
  });
});

describe("Copilot conversations", () => {
  const CLI_ID = "4d1c0f6a-0000-4000-8000-000000000001";
  const env = () => homes({ COPILOT_HOME: (root) => path.join(root, "copilot") }, { COPILOT_HOME: (root) => path.join(root, "copilot-private") });

  it("reads a Copilot CLI conversation and copies its folder into the add-in's private Copilot folder", async () => {
    const h = env();
    const dir = path.join(h.person, "copilot", "session-state", CLI_ID);
    write(path.join(dir, "workspace.yaml"), "id: " + CLI_ID + "\ncwd: /work/app\nclient_name: vscode\nname: Tidy imports\nupdated_at: 2026-01-01T00:10:00.000Z\n");
    write(path.join(dir, "events.jsonl"), jsonl([
      { type: "session.start", data: { sessionId: CLI_ID }, timestamp: "2026-01-01T00:00:00Z" },
      { type: "user.message", data: { content: "tidy the imports", transformedContent: "<time>tidy the imports" }, timestamp: "2026-01-01T00:00:01Z" },
      { type: "assistant.message", data: { content: "", toolRequests: [{ toolCallId: "k1", name: "edit", toolTitle: "Edit file", arguments: { path: "a.ts" } }] } },
      { type: "tool.execution_complete", data: { toolCallId: "k1", success: true, result: { content: "edited" } } },
      { type: "assistant.message", data: { content: "Imports are tidy." }, timestamp: "2026-01-01T00:00:09Z" },
    ]));
    write(path.join(dir, "files", "notes.txt"), "kept");
    expect(await listCopilotSessions(h.homes.source)).toEqual([expect.objectContaining({ id: CLI_ID, title: "Tidy imports", origin: "VS Code", cwd: "/work/app" })]);
    const imported = await importCopilotSession(h.homes, CLI_ID);
    expect(imported.providerSessionId).toBe(CLI_ID);
    expect(imported.messages).toEqual([
      { kind: "user", text: "tidy the imports" },
      { kind: "tool_group", text: "", tools: [{ name: "Edit file", args: '{"path":"a.ts"}', result: "edited", state: "ok", error: null }] },
      { kind: "assistant", text: "Imports are tidy." },
    ]);
    expect(fs.readFileSync(path.join(h.product, "copilot-private", "session-state", CLI_ID, "files", "notes.txt"), "utf8")).toBe("kept");
  });

  describe("chats of the Copilot Chat extension in VS Code", () => {
    const SESSION = "b611f895-bc56-4a8c-9a4e-54aaca2803d5";
    const chatLog = [
      { kind: 0, v: { version: 3, creationDate: 1_700_000_000_000, sessionId: SESSION, requests: [] } },
      { kind: 2, k: ["requests"], v: [{ requestId: "r1", message: { text: "Explain closures", parts: [] }, response: [], timestamp: 1_700_000_001_000 }] },
      { kind: 2, k: ["requests", 0, "response"], v: [{ value: "A closure " }, { kind: "inlineReference", inlineReference: { path: "/work/app/a.ts" } }, { kind: "markdownContent", content: { value: " keeps its scope." } }] },
      { kind: 2, k: ["requests"], v: [{ requestId: "r2", message: "Thanks", response: [{ kind: "toolInvocationSerialized", toolId: "copilot_readFile", toolCallId: "tc1", invocationMessage: "Reading a.ts", pastTenseMessage: { value: "Read a.ts" } }, { value: "You're welcome." }], timestamp: 1_700_000_005_000 }] },
      { kind: 1, k: ["customTitle"], v: "Closures" },
    ];
    const setup = () => {
      const h = env();
      const user = process.platform === "darwin" ? path.join(h.person, "Library", "Application Support", "Code", "User") : path.join(h.person, "Code", "User");
      const storage = path.join(user, "workspaceStorage", "hash1");
      write(path.join(storage, "chatSessions", `${SESSION}.jsonl`), jsonl(chatLog));
      write(path.join(storage, "workspace.json"), JSON.stringify({ folder: "file:///work/app" }));
      write(path.join(storage, "chatSessions", "a0000000-0000-4000-8000-000000000002.jsonl"), jsonl([{ kind: 0, v: { version: 3, requests: [] } }]));
      return h;
    };

    it("lists the chats that have questions, not the panels that were opened and left empty", async () => {
      const list = await listCopilotSessions(setup().homes.source);
      expect(list).toEqual([expect.objectContaining({ id: `vscode-chat-${SESSION}`, title: "Closures", origin: "VS Code", messageCount: 2 })]);
    });

    it("writes the chat as a Copilot CLI conversation, once, so the CLI can resume it with the same context", async () => {
      const h = setup();
      const imported = await importCopilotSession(h.homes, `vscode-chat-${SESSION}`, "1.0.112");
      expect(imported.messages).toEqual([
        { kind: "user", text: "Explain closures" },
        { kind: "assistant", text: "A closure `a.ts` keeps its scope." },
        { kind: "user", text: "Thanks" },
        { kind: "tool_group", text: "", tools: [{ name: "copilot_readFile", args: "Reading a.ts", result: "Read a.ts", state: "ok", error: null }] },
        { kind: "assistant", text: "You're welcome." },
      ]);
      expect(imported.providerSessionId).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
      const dir = path.join(h.product, "copilot-private", "session-state", imported.providerSessionId);
      const events = fs.readFileSync(path.join(dir, "events.jsonl"), "utf8").trim().split("\n").map((line) => JSON.parse(line));
      expect(events.map((event) => event.type)).toEqual(["session.start", ...Array(2).fill(["user.message", "assistant.turn_start", "assistant.message", "assistant.turn_end"]).flat()]);
      expect(events.filter((event) => event.type === "user.message").map((event) => event.data.content)).toEqual(["Explain closures", "Thanks"]);
      expect(events.filter((event) => event.type === "assistant.message").map((event) => event.data.content)).toEqual(["A closure `a.ts` keeps its scope.", "You're welcome."]);
      expect(events.every((event, index) => event.parentId === (index === 0 ? null : events[index - 1].id))).toBe(true);
      expect(events[0].data).toMatchObject({ copilotVersion: "1.0.112", context: { cwd: ensureWorkspaceDir() } });
      expect(parse(fs.readFileSync(path.join(dir, "workspace.yaml"), "utf8"))).toMatchObject({ id: imported.providerSessionId, name: "Closures", cwd: ensureWorkspaceDir() });
      // The same chat is the same conversation, and what happened in it since is not overwritten.
      fs.appendFileSync(path.join(dir, "events.jsonl"), "{}\n");
      const again = await importCopilotSession(h.homes, `vscode-chat-${SESSION}`);
      expect(again.providerSessionId).toBe(imported.providerSessionId);
      expect(fs.readFileSync(path.join(dir, "events.jsonl"), "utf8").endsWith("{}\n")).toBe(true);
    });
  });
});

describe("History keeps an imported conversation once", () => {
  const session = (): StoredSession => ({ id: "claude:abc", title: "Imported", provider: "claude", host: "excel", providerSessionId: "abc", identityVersion: 1, startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "hello" }], changes: [] });
  it("adds it, leaves it alone the second time, and adds it back after it was deleted", () => {
    const dir = temp("history-import-"), store = new HistoryStore(path.join(dir, "history.db"));
    try {
      expect(store.saveImported("excel", session())).toEqual({ id: "claude:abc", existing: false });
      const continued = { ...session(), messages: [{ kind: "user" as const, text: "hello" }, { kind: "assistant" as const, text: "continued here" }] };
      store.save("excel", { session: continued, baseRevision: 1, generation: 0, operationId: "continue" });
      expect(store.saveImported("excel", session())).toEqual({ id: "claude:abc", existing: true });
      expect(store.get("excel", "claude:abc")!.session.messages).toHaveLength(2);
      store.delete("excel", "claude:abc");
      expect(store.saveImported("excel", session())).toEqual({ id: "claude:abc", existing: false });
      expect(store.get("excel", "claude:abc")!.session.messages).toHaveLength(1);
    } finally { store.close(); }
  });
});
