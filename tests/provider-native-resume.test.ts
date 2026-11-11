import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { ProviderEvent } from "../protocol/src/index.js";
import type { StartSessionContext } from "../providers/src/types.js";

const rpc = vi.hoisted(() => ({
  calls: [] as Array<{ method: string; params: any }>,
  launches: [] as any[],
  reply: (async (_method: string, _params: any, _options: any): Promise<any> => ({})),
}));
vi.mock("../providers/src/jsonrpc.js", () => ({
  JsonRpcPeer: class {
    alive = true;
    stderrTail: string[] = [];
    constructor(private options: any) { rpc.launches.push(options); }
    async request(method: string, params: any) { rpc.calls.push({ method, params }); return rpc.reply(method, params, this.options); }
    notify() {}
    onExit() {}
    async close() { this.alive = false; }
  },
}));
import * as base from "../providers/src/base.js";
import { CopilotAdapter } from "../providers/src/copilot.js";
import { OpencodeAdapter } from "../providers/src/opencode.js";
import { CodexAdapter } from "../providers/src/codex.js";
import { ClaudeAdapter } from "../providers/src/claude.js";
vi.mock("../providers/src/claude-process.js", () => ({
  ClaudeProcess: class {
    async request() { return {}; }
    async close() {}
  },
}));

const runtime = { bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } };
const nativeId = "native-session";
const context = (): StartSessionContext => ({ workbookConnectionId: "test-workbook", sessionId: "bridge-session",
  resumeSessionId: nativeId, signal: new AbortController().signal,
  mcp: { command: process.execPath, args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" } });
let dir = "";
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), "native-resume-"));
  vi.stubEnv("EXCEL_AI_HOME", dir);
  vi.stubEnv("COPILOT_HOME", dir);
  rpc.calls = [];
  rpc.launches = [];
  rpc.reply = async () => ({});
  vi.spyOn(OpencodeAdapter.prototype, "detect").mockResolvedValue({ provider: "opencode", installed: true, version: "1.19.3", status: "ready", hint: null, capabilities: { streaming: true, cancel: true, sessionResume: true, mcp: true, acp: true } });
});
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllEnvs(); fs.rmSync(dir, { recursive: true, force: true }); });
async function turn(adapter: CopilotAdapter | OpencodeAdapter | CodexAdapter, internalId: string) {
  const events: ProviderEvent[] = [];
  for await (const event of adapter.sendTurn({ providerSessionId: internalId, requestId: "request", prompt: "new request",
    selection: null, signal: new AbortController().signal })) events.push(event);
  return events;
}

describe("provider-native web search", () => {
  it.each([false, true])("configures Codex hosted search without adding a search MCP (%j)", async enabled => {
    rpc.reply = async method => method === "thread/resume" ? { thread: { id: nativeId } } : {};
    const adapter = new CodexAdapter(runtime);
    const ctx = context();
    if (enabled) ctx.mcp.env.EXCEL_WEB_SEARCH = "1";
    try {
      await adapter.startSession(ctx);
      const config = rpc.calls.find(call => call.method === "thread/resume")!.params.config;
      expect(config.web_search).toBe(enabled ? "live" : "disabled");
      const instructions = rpc.calls.find(call => call.method === "thread/resume")!.params.developerInstructions;
      expect(instructions).toContain("supersedes Excel-only and no-web restrictions");
      expect(instructions).toContain(enabled ? "Native web search: enabled" : "Native web search: disabled");
      expect(rpc.launches.at(-1).args).toContain(`web_search="${enabled ? "live" : "disabled"}"`);
      expect(Object.keys(config.mcp_servers)).toEqual(["excel"]);
    } finally { await adapter.dispose(); }
  });

  it.each([false, true])("exposes only Copilot native web tools and rejects unrelated permission requests (%j)", async enabled => {
    const adapter = new CopilotAdapter(runtime);
    const ctx = context();
    if (enabled) ctx.mcp.env.EXCEL_WEB_SEARCH = "1";
    try {
      await adapter.startSession(ctx);
      const launch = rpc.launches.at(-1)!;
      expect(launch.args.includes("web_search")).toBe(enabled);
      expect(launch.args.includes("web_fetch")).toBe(enabled);
      expect(launch.args).not.toContain("github-mcp-server-web_search");
      expect(launch.args).not.toContain("github-mcp-server(web_search)");
      expect(launch.args.includes("--disable-builtin-mcps")).toBe(!enabled);
      const options = [{ optionId: "yes", kind: "allow_once" }, { optionId: "no", kind: "reject_once" }];
      expect(launch.onRequest("session/request_permission", { toolCall: { title: "web_search" }, options }))
        .toEqual({ outcome: { outcome: "selected", optionId: enabled ? "yes" : "no" } });
      expect(launch.onRequest("session/request_permission", { toolCall: { title: "shell" }, options }))
        .toEqual({ outcome: { outcome: "selected", optionId: "no" } });
      expect(launch.onRequest("session/request_permission", { toolCall: { title: "github-mcp-server-web_search" }, options }))
        .toEqual({ outcome: { outcome: "selected", optionId: "no" } });
    } finally { await adapter.dispose(); }
  });

  it("keeps OpenCode external search disabled", async () => {
    const adapter = new OpencodeAdapter(runtime);
    try {
      await adapter.startSession(context());
      const file = rpc.launches.at(-1)!.env.OPENCODE_CONFIG;
      expect(JSON.parse(fs.readFileSync(file, "utf8")).permission.websearch).toBe("deny");
    } finally { await adapter.dispose(); }
  });
});

describe("optional ACP add-in instructions", () => {
  it.each([CopilotAdapter, OpencodeAdapter])("sends a clean first prompt without instructions and still supports a custom prompt (%j)", async Adapter => {
    vi.spyOn(base, "execCli").mockResolvedValue({ code: 0, stdout: "{}", stderr: "" });
    rpc.reply = async method => method === "session/new" ? { sessionId: nativeId }
      : method === "session/prompt" ? { stopReason: "end_turn" } : {};
    for (const systemPrompt of [undefined, "", " ", "Custom instructions"]) {
      const adapter = new Adapter(runtime);
      try {
        const internalId = await adapter.startSession({ ...context(), resumeSessionId: undefined });
        for await (const _event of adapter.sendTurn({ providerSessionId: internalId, requestId: "request",
          prompt: "Hello", selection: null, signal: new AbortController().signal, systemPrompt })) {}
        const prompt = rpc.calls.filter(call => call.method === "session/prompt").at(-1)!.params.prompt;
        expect(prompt).toEqual([{ type: "text", text: systemPrompt?.trim() ? `${systemPrompt}\n\n---\n\nHello` : "Hello" }]);
      } finally { await adapter.dispose(); }
    }
  });
});

describe("Copilot automatic session naming", () => {
  it("places the original user request before injected context and custom instructions", async () => {
    vi.spyOn(base, "execCli").mockResolvedValue({ code: 0, stdout: "{}", stderr: "" });
    rpc.reply = async method => method === "session/new" ? { sessionId: nativeId }
      : method === "session/prompt" ? { stopReason: "end_turn" } : {};
    const adapter = new CopilotAdapter(runtime);
    try {
      const internalId = await adapter.startSession({ ...context(), resumeSessionId: undefined });
      for await (const _event of adapter.sendTurn({ providerSessionId: internalId, requestId: "request",
        userPrompt: "Tạo báo cáo PHP", prompt: "<saved_reference_files>files</saved_reference_files>\nTạo báo cáo PHP",
        selection: null, signal: new AbortController().signal, systemPrompt: "Custom instructions" })) {}
      const text = rpc.calls.filter(call => call.method === "session/prompt").at(-1)!.params.prompt[0].text;
      expect(text).toMatch(/^Tạo báo cáo PHP\n\n<addin_context>/);
      expect(text).toContain("Custom instructions");
      expect(text).toContain("<saved_reference_files>files</saved_reference_files>");
    } finally { await adapter.dispose(); }
  });
});

describe("each provider resumes its native session independently", () => {
  it("Copilot loads the old ACP session and subtracts the original pre-resume accounting", async () => {
    const file = path.join(dir, "session-state", nativeId, "events.jsonl");
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(path.join(path.dirname(file), "workspace.yaml"), 'name: Native Copilot Title\n');
    const record = (id: string, inputTokens: number, outputTokens: number, nano: number) => JSON.stringify({
      type: "session.usage_record", id, data: { usage: { inputTokens, outputTokens, cacheReadTokens: 30,
        cacheWriteTokens: 10, copilotUsage: { totalNanoAiu: nano }, aiCreditsStatus: "complete" } },
    }) + "\n";
    fs.writeFileSync(file, record("before-resume", 100, 20, 1234567890));
    rpc.reply = async method => {
      if (method === "session/prompt") {
        fs.appendFileSync(file, record("after-resume", 50, 5, 11111111));
        return { stopReason: "end_turn", usage: { inputTokens: 150, outputTokens: 25, cachedReadTokens: 60, totalTokens: 175 } };
      }
      return {};
    };
    const adapter = new CopilotAdapter(runtime);
    const internalId = await adapter.startSession(context());
    try {
      const events = await turn(adapter, internalId);
      expect(rpc.calls.some(c => c.method === "session/new")).toBe(false);
      expect(rpc.calls.find(c => c.method === "session/load")?.params.sessionId).toBe(nativeId);
      expect(events.find(e => e.type === "session_info")).toMatchObject({ providerSessionId: nativeId, title: "Native Copilot Title" });
      expect(events.find(e => e.type === "usage_update")).toMatchObject({ usage: {
        totalTokens: 55, copilotCreditNanoAiu: 11111111, copilotCredits: 0.011111111,
      } });
    } finally { await adapter.dispose(); }
  });

  it.each([false, true])("OpenCode loads its ACP session and uses exported totals including cache and reasoning (v2=%j)", async v2 => {
    if (v2) vi.mocked(OpencodeAdapter.prototype.detect).mockResolvedValue({ provider: "opencode", installed: true, version: "2.0.32", status: "ready", hint: null, capabilities: { streaming: true, cancel: true, sessionResume: true, mcp: true, acp: true } });
    let exports = 0;
    vi.spyOn(base, "execCli").mockImplementation(async (_command, args) => {
      if (v2 ? args[0] === "session" && args[1] === "export" : args[0] === "export") {
        exports++;
        const tokens = exports === 1
          ? { input: 100, output: 20, reasoning: 5, cache: { read: 30, write: 10 } }
          : { input: 200, output: 50, reasoning: 8, cache: { read: 60, write: 20 } };
        return { code: 0, stdout: JSON.stringify({ info: { id: nativeId, title: "Native OpenCode Title", tokens }, ...(v2 ? { messages: [{ id: "msg", type: "assistant", tokens }] } : {}) }), stderr: "" };
      }
      return { code: 0, stdout: "[]", stderr: "" };
    });
    rpc.reply = async method => method === "session/prompt"
      ? { stopReason: "end_turn", usage: { inputTokens: 1, outputTokens: 1, totalTokens: 2 } } : {};
    const adapter = new OpencodeAdapter(runtime);
    const ctx = context();
    if (v2) ctx.mcp.toolNames = ["execute_office_js"];
    const internalId = await adapter.startSession(ctx);
    try {
      const events = await turn(adapter, internalId);
      expect(rpc.calls.some(c => c.method === "session/new")).toBe(false);
      const load = rpc.calls.find(c => c.method === "session/load")!;
      expect(load.params.sessionId).toBe(nativeId);
      expect(load.params.mcpServers).toHaveLength(1);
      if (v2) {
        const config = JSON.parse(fs.readFileSync(rpc.launches.at(-1)!.env.OPENCODE_CONFIG, "utf8"));
        expect(config.permissions).toContainEqual({ action: "websearch", resource: "*", effect: "deny" });
        expect(config.permissions).toContainEqual({ action: "execute", resource: "*", effect: "allow" });
        expect(config.permissions).toContainEqual({ action: "excel_*", resource: "*", effect: "allow" });
        expect(config.permissions).toContainEqual({ action: "*", resource: "*", effect: "ask" });
        expect(config.permission).toBeUndefined();
      }
      expect(events.filter(e => e.type === "session_info").at(-1)).toMatchObject({ providerSessionId: nativeId, title: "Native OpenCode Title" });
      expect(events.find(e => e.type === "usage_update")).toMatchObject({ usage: {
        totalTokens: 173, inputTokens: 140, outputTokens: 33, reasoningTokens: 3, cachedInputTokens: 30, cacheWriteTokens: 10,
      } });
    } finally { await adapter.dispose(); }
  });

  it("keeps workbook add-in OpenCode V2 sessions on their existing configuration and export path", async () => {
    const adapter = new OpencodeAdapter(runtime);
    const detected = await adapter.detect();
    vi.mocked(adapter.detect).mockResolvedValue({ ...detected, version: "2.0.32" });
    const exportCli = vi.spyOn(base, "execCli").mockResolvedValue({ code: 0, stderr: "", stdout: JSON.stringify({
      info: { id: nativeId, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } } },
    }) });
    const ctx = context();
    ctx.requestUserInput = async () => ({ cancelled: true, answers: {} });
    try {
      await adapter.startSession(ctx);
      const launch = rpc.launches.at(-1)!;
      const config = JSON.parse(fs.readFileSync(launch.env.OPENCODE_CONFIG, "utf8"));
      expect(config.permission).toMatchObject({ "*": "ask", websearch: "deny", "excel_*": "allow" });
      expect(config.permissions).toBeUndefined();
      expect(exportCli).toHaveBeenCalledWith(expect.any(String), ["export", nativeId, "--pure"], expect.any(Object), 10_000);
      expect(launch.args).toEqual(["acp"]);
    } finally { await adapter.dispose(); }
  });

  it("Codex resumes a persistent thread and uses the rollout odometer as its baseline", async () => {
    const file = path.join(dir, "rollout.jsonl");
    fs.writeFileSync(file, JSON.stringify({ type: "event_msg", payload: { type: "token_count", info: {
      total_token_usage: { input_tokens: 100, output_tokens: 20, total_tokens: 120, cached_input_tokens: 30 },
    } } }) + "\n");
    rpc.reply = async (method, _params, options) => {
      if (method === "thread/resume") return { thread: { id: nativeId, name: "Native Codex Title", path: file } };
      if (method === "thread/read") return { thread: { name: "Native Codex Title" } };
      if (method === "turn/start") {
        queueMicrotask(() => {
          options.onNotification("item/started", { threadId: nativeId, item: { type: "webSearch", id: "search", query: "Google", action: null } });
          options.onNotification("item/completed", { threadId: nativeId, item: { type: "webSearch", id: "search", query: "Google", action: null } });
          options.onNotification("thread/tokenUsage/updated", { threadId: nativeId, turnId: "turn", tokenUsage: {
            total: { inputTokens: 160, outputTokens: 30, totalTokens: 190, cachedInputTokens: 40 },
          } });
          options.onNotification("turn/completed", { threadId: nativeId, turn: { id: "turn", status: "completed" } });
        });
        return { turn: { id: "turn", status: "inProgress" } };
      }
      return {};
    };
    const adapter = new CodexAdapter(runtime);
    const internalId = await adapter.startSession(context());
    try {
      const events = await turn(adapter, internalId);
      expect(rpc.calls.some(c => c.method === "thread/start")).toBe(false);
      expect(rpc.calls.find(c => c.method === "thread/resume")?.params.threadId).toBe(nativeId);
      expect(events.find(e => e.type === "session_info")).toMatchObject({ providerSessionId: nativeId, title: "Native Codex Title" });
      expect(events.find(e => e.type === "tool_started")).toMatchObject({ tool: "web_search", callId: "search" });
      expect(events.find(e => e.type === "tool_finished")).toMatchObject({ tool: "web_search", ok: true });
      expect(events.find(e => e.type === "usage_update")).toMatchObject({ usage: { totalTokens: 70, cachedInputTokens: 10, source: "codex_thread" } });
    } finally { await adapter.dispose(); }
  });

  it("Claude keeps the native conversation ID when opening a saved conversation", async () => {
    const adapter = new ClaudeAdapter(runtime);
    const internalId = await adapter.startSession(context());
    try {
      const sessions = (adapter as unknown as { sessions: Map<string, { conversationId: string; resumable: boolean }> }).sessions;
      expect(sessions.get(internalId)).toMatchObject({ conversationId: nativeId, resumable: true });
    } finally { await adapter.dispose(); }
  });

  it.each(["copilot", "opencode", "codex"])("%s reports a failed native resume without creating a replacement", async provider => {
    rpc.reply = async method => {
      if (method === "session/load" || method === "thread/resume") throw new Error("native session no longer exists");
      return {};
    };
    const adapter = provider === "copilot" ? new CopilotAdapter(runtime) : provider === "opencode" ? new OpencodeAdapter(runtime) : new CodexAdapter(runtime);
    await expect(adapter.startSession(context())).rejects.toThrow("no longer exists");
    expect(rpc.calls.some(c => c.method === "session/new" || c.method === "thread/start")).toBe(false);
  });
});

it("O03 rejects a failed model selection instead of silently using the default", async () => {
  rpc.reply = async method => { if (method === "session/set_config_option") throw Error("Unknown model"); return { sessionId: nativeId }; };
  const adapter = new OpencodeAdapter(runtime);
  try { await expect(adapter.startSession({ ...context(), resumeSessionId: undefined, model: "provider/missing" })).rejects.toThrow("Unknown model"); }
  finally { await adapter.dispose(); }
});

it("O04 reports an empty OpenCode response as an error", async () => {
  rpc.reply = async method => method === "session/prompt" ? { stopReason: "end_turn" } : { sessionId: nativeId };
  vi.spyOn(base, "execCli").mockResolvedValue({ code: 1, stdout: "", stderr: "" });
  const adapter = new OpencodeAdapter(runtime);
  try {
    const id = await adapter.startSession(context());
    const events = await turn(adapter, id);
    expect(events).toContainEqual(expect.objectContaining({ type: "error", code: "empty_response" }));
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "error" });
  } finally { await adapter.dispose(); }
});

it("O05 gives a silent but healthy OpenCode request ten minutes before releasing the pane", async () => {
  vi.useFakeTimers();
  rpc.reply = async method => method === "session/prompt" ? new Promise(() => {}) : { sessionId: nativeId };
  vi.spyOn(base, "execCli").mockResolvedValue({ code: 1, stdout: "", stderr: "" });
  const adapter = new OpencodeAdapter(runtime);
  try {
    const id = await adapter.startSession({ ...context(), resumeSessionId: undefined });
    let done = false;
    const result = turn(adapter, id).then(events => { done = true; return events; });
    await vi.advanceTimersByTimeAsync(120_001);
    expect(done).toBe(false);
    await vi.advanceTimersByTimeAsync(480_000);
    const events = await result;
    expect(events).toContainEqual(expect.objectContaining({ type: "error", code: "provider_timeout", message: expect.stringContaining("600 seconds") }));
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "error" });
  } finally { vi.useRealTimers(); await adapter.dispose(); }
});

it("O05 stops a request that OpenCode keeps retrying after two minutes", async () => {
  vi.useFakeTimers();
  rpc.reply = async (method, _params, options) => {
    if (method !== "session/prompt") return { sessionId: nativeId };
    options.onNotification("session/update", { sessionId: nativeId, update: { sessionUpdate: "session_info_update", _meta: { "opencode/retry": { attempt: 2, error: { status: 503 } } } } });
    return new Promise(() => {});
  };
  vi.spyOn(base, "execCli").mockResolvedValue({ code: 1, stdout: "", stderr: "" });
  const adapter = new OpencodeAdapter(runtime);
  try {
    const id = await adapter.startSession({ ...context(), resumeSessionId: undefined });
    const result = turn(adapter, id);
    await vi.advanceTimersByTimeAsync(120_001);
    const events = await result;
    expect(events).toContainEqual(expect.objectContaining({ type: "provider_progress", state: "retrying" }));
    expect(events).toContainEqual(expect.objectContaining({ type: "error", code: "provider_timeout", message: expect.stringContaining("120 seconds") }));
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "error" });
  } finally { vi.useRealTimers(); await adapter.dispose(); }
});

it("O05 keeps streaming work running after the initial-response deadline", async () => {
  vi.useFakeTimers();
  let finish!: (value: unknown) => void;
  rpc.reply = async (method, _params, options) => {
    if (method !== "session/prompt") return { sessionId: nativeId };
    options.onNotification("session/update", { sessionId: nativeId, update: { sessionUpdate: "agent_message_chunk", content: { type: "text", text: "Answer" } } });
    return new Promise(resolve => { finish = resolve; });
  };
  vi.spyOn(base, "execCli").mockResolvedValue({ code: 1, stdout: "", stderr: "" });
  const adapter = new OpencodeAdapter(runtime);
  try {
    const id = await adapter.startSession({ ...context(), resumeSessionId: undefined });
    const result = turn(adapter, id);
    await vi.advanceTimersByTimeAsync(120_001);
    finish({ stopReason: "end_turn" });
    const events = await result;
    expect(events.some(event => event.type === "error")).toBe(false);
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "completed" });
  } finally { vi.useRealTimers(); await adapter.dispose(); }
});

it("O05 cancels a silent request without waiting for the CLI prompt response", async () => {
  rpc.reply = async method => method === "session/prompt" ? new Promise(() => {}) : { sessionId: nativeId };
  const adapter = new OpencodeAdapter(runtime), controller = new AbortController();
  try {
    const id = await adapter.startSession({ ...context(), resumeSessionId: undefined });
    const events: ProviderEvent[] = [];
    const running = (async () => { for await (const event of adapter.sendTurn({ providerSessionId: id, requestId: "r", prompt: "Hello", selection: null, signal: controller.signal })) {
      events.push(event); if (event.type === "provider_progress") controller.abort();
    } })();
    await running;
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "cancelled" });
  } finally { await adapter.dispose(); }
});
