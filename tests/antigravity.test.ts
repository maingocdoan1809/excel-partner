import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AntigravityAdapter } from "../providers/src/antigravity/adapter.js";
import { ProviderLoginError } from "../providers/src/login.js";
import type { ProviderEvent } from "../protocol/src/index.js";
import type { SendTurnInput, StartSessionContext } from "../providers/src/types.js";

const fake = vi.hoisted(() => ({ peers: [] as any[], resume: true, fail: "", prompt: undefined as undefined | ((peer: any, params: any) => Promise<any>) }));
vi.mock("../providers/src/jsonrpc.js", () => ({ JsonRpcPeer: class {
  alive = true;
  exit?: () => void;
  request = vi.fn(async (method: string, params: any) => {
    if (method === "initialize") return { agentCapabilities: { sessionCapabilities: fake.resume ? { resume: {} } : {} } };
    if (["session/new", "session/resume", "session/load"].includes(method)) {
      if (fake.fail) throw new Error(fake.fail);
      return { sessionId: "native-1", configOptions: [
        { id: "model", category: "model", type: "select", currentValue: "gemini-pro", options: [{ value: "gemini-pro", name: "Gemini Pro" }, { value: "gemini-flash", name: "Gemini Flash" }] },
        { id: "thinking", category: "thought_level", type: "select", currentValue: "medium", options: [{ value: "medium" }, { value: "high" }] },
      ], modes: { availableModes: [{ id: "default" }], currentModeId: "auto" } };
    }
    if (method === "session/prompt") {
      if (fake.prompt) return fake.prompt(this, params);
      this.note({ sessionUpdate: "agent_message_chunk", content: { type: "text", text: "Hello" } });
      return { stopReason: "end_turn" };
    }
    return {};
  });
  notify = vi.fn();
  close = vi.fn(async () => { this.alive = false; this.exit?.(); });
  constructor(readonly options: any) { fake.peers.push(this); }
  onExit(callback: () => void) { this.exit = callback; }
  note(update: unknown) { this.options.onNotification?.("session/update", { sessionId: "native-1", update }); }
} }));
const adapters: AntigravityAdapter[] = [];
const ctx = (extra: Partial<StartSessionContext> = {}): StartSessionContext => ({
  sessionId: "chat", workbookConnectionId: "file", signal: new AbortController().signal, systemPrompt: "Office instructions",
  mcp: { command: "node", args: ["mcp.js"], env: { EXCEL_SESSION_TOKEN: "test-token", OFFICE_HOST: "word" }, bridgeUrl: "", sessionId: "chat", token: "test-token" }, ...extra,
});
const input = (id: string, extra: Partial<SendTurnInput> = {}): SendTurnInput => ({ providerSessionId: id, requestId: "r", prompt: "hello", selection: null, signal: new AbortController().signal, ...extra });
async function collect(events: AsyncIterable<ProviderEvent>) { const result: ProviderEvent[] = []; for await (const event of events) result.push(event); return result; }
function adapter() {
  const result = new AntigravityAdapter({ env: { ANTIGRAVITY_BIN: "agy-test", ANTIGRAVITY_GEMINI_DIR: "private-state" }, bridgeUrl: "", mcp: { command: "node", args: [] }, logger: { debug() {}, warn() {} } });
  adapters.push(result); return result;
}
beforeEach(() => { fake.peers.length = 0; fake.resume = true; fake.fail = ""; fake.prompt = undefined; });
afterEach(async () => { await Promise.all(adapters.splice(0).map(a => a.dispose())); });

describe("Antigravity through the existing provider contract", () => {
  it("registers the supplied MCP endpoint and keeps one runtime for successive turns", async () => {
    const a = adapter(), context = ctx(), id = await a.startSession(context), peer = fake.peers[0];
    expect(peer.options).toMatchObject({ command: "agy-test", args: ["acp", "--disable-telemetry", "--gemini-dir", "private-state"], versionField: true });
    expect(peer.request).toHaveBeenCalledWith("session/new", expect.objectContaining({ mcpServers: [{ type: "stdio", name: "excel", command: "node", args: ["mcp.js"], env: [{ name: "EXCEL_SESSION_TOKEN", value: "test-token" }, { name: "OFFICE_HOST", value: "word" }] }] }), 30_000);
    expect(peer.request).toHaveBeenCalledWith("session/set_mode", { sessionId: "native-1", modeId: "default" }, 15_000);
    expect(await collect(a.sendTurn(input(id)))).toContainEqual({ type: "message_completed", requestId: "r", text: "Hello" });
    await collect(a.sendTurn(input(id, { requestId: "r2", model: "gemini-flash", effort: "high", images: [{ mediaType: "image/png", data: "aGVsbG8=" }] })));
    expect(fake.peers).toHaveLength(1);
    const prompts = peer.request.mock.calls.filter(([method]: string[]) => method === "session/prompt");
    expect(prompts[0][1].prompt[0].text).toContain("Office instructions");
    expect(prompts[1][1].prompt).toEqual([{ type: "text", text: "hello" }, { type: "image", mimeType: "image/png", data: "aGVsbG8=" }]);
    expect(peer.request).toHaveBeenCalledWith("session/set_config_option", { sessionId: "native-1", configId: "model", value: "gemini-flash" }, 15_000);
    expect(peer.request).toHaveBeenCalledWith("session/set_config_option", { sessionId: "native-1", configId: "thinking", value: "high" }, 15_000);
  });
  it("resumes without replaying history, and never creates a replacement when resume fails", async () => {
    const a = adapter();
    await a.startSession(ctx({ resumeSessionId: "native-1" }));
    expect(fake.peers[0].request.mock.calls.some(([m]: string[]) => m === "session/resume")).toBe(true);
    fake.fail = "Session was deleted";
    await expect(a.startSession(ctx({ resumeSessionId: "missing" }))).rejects.toThrow("deleted");
    expect(fake.peers[1].request.mock.calls.some(([m]: string[]) => m === "session/new")).toBe(false);
    expect(fake.peers[1].close).toHaveBeenCalled();
  });
  it("reports authentication errors and discovers models without sending an inference prompt", async () => {
    const a = adapter();
    expect(await a.listModels()).toMatchObject({ defaultModelId: "gemini-pro", models: [{ id: "gemini-pro", name: "Gemini Pro" }, { id: "gemini-flash", name: "Gemini Flash" }] });
    expect(fake.peers[0].request.mock.calls.some(([m]: string[]) => m === "session/prompt")).toBe(false);
    expect(fake.peers[0].close).toHaveBeenCalled();
    fake.fail = "Authentication required";
    await expect(a.startSession(ctx())).rejects.toBeInstanceOf(ProviderLoginError);
  });
  it("routes permissions through existing Office and native tool approval", async () => {
    const authorizeTool = vi.fn(async () => false), a = adapter();
    await a.startSession(ctx({ authorizeTool }));
    const options = [{ kind: "allow_once", optionId: "yes" }, { kind: "reject_once", optionId: "no" }];
    const request = fake.peers[0].options.onRequest;
    expect(await request("session/request_permission", { toolCall: { title: "excel/read_range" }, options })).toMatchObject({ outcome: { optionId: "yes" } });
    expect(authorizeTool).not.toHaveBeenCalled();
    expect(await request("session/request_permission", { toolCall: { title: "shell" }, options })).toMatchObject({ outcome: { optionId: "no" } });
    expect(authorizeTool).toHaveBeenCalledOnce();
    await expect(request("fs/read_text_file", {})).rejects.toThrow("unsupported");
  });
  it("streams reasoning and tool errors with exactly one final event", async () => {
    fake.prompt = async peer => {
      peer.note({ sessionUpdate: "agent_thought_chunk", content: { type: "text", text: "Thinking" } });
      peer.note({ sessionUpdate: "tool_call", toolCallId: "t", title: "excel/read_range", rawInput: { address: "A1" } });
      peer.note({ sessionUpdate: "tool_call_update", toolCallId: "t", status: "failed", rawOutput: "Denied" });
      peer.note({ sessionUpdate: "tool_call_update", toolCallId: "t", status: "failed" });
      return { stopReason: "end_turn" };
    };
    const a = adapter(), id = await a.startSession(ctx()), events = await collect(a.sendTurn(input(id)));
    expect(events).toContainEqual({ type: "thinking_delta", requestId: "r", text: "Thinking" });
    expect(events.filter(e => e.type === "tool_finished")).toEqual([{ type: "tool_finished", requestId: "r", callId: "t", tool: "read_range", ok: false, error: "Denied" }]);
    expect(events.filter(e => e.type === "turn_finished")).toHaveLength(1);
  });
  it("waits for cancel confirmation before reusing the same runtime", async () => {
    let release!: (value: unknown) => void;
    fake.prompt = () => new Promise(resolve => { release = resolve; });
    const a = adapter(), id = await a.startSession(ctx()), events: ProviderEvent[] = [];
    let finished = false;
    const running = (async () => { for await (const e of a.sendTurn(input(id))) events.push(e); finished = true; })();
    await vi.waitFor(() => expect(release).toBeTypeOf("function"));
    await a.cancelTurn(id);
    expect(fake.peers[0].notify).toHaveBeenCalledWith("session/cancel", { sessionId: "native-1" });
    expect(finished).toBe(false);
    release({ stopReason: "cancelled" }); await running;
    expect(events.filter(e => e.type === "turn_finished")).toEqual([{ type: "turn_finished", requestId: "r", reason: "cancelled" }]);
    fake.prompt = undefined;
    expect(await collect(a.sendTurn(input(id)))).toContainEqual({ type: "turn_finished", requestId: "r", reason: "completed" });
  });
  it("extracts Google's sign-in URL without treating unrelated URLs as credentials", () => {
    expect(AntigravityAdapter.login.find("Open https://accounts.google.com/o/oauth2/auth?client_id=test&state=x\nEnter code:")).toEqual({ uri: "https://accounts.google.com/o/oauth2/auth?client_id=test&state=x", code: null });
    expect(AntigravityAdapter.login.find("https://example.com/login")).toBeNull();
  });
  it("releases turn state even when the consumer stops after the first event", async () => {
    let release!: (value: unknown) => void;
    fake.prompt = () => new Promise(resolve => { release = resolve; });
    const a = adapter(), id = await a.startSession(ctx()), iterator = a.sendTurn(input(id))[Symbol.asyncIterator]();
    await iterator.next();
    const returned = iterator.return!();
    release({ stopReason: "cancelled" }); await returned;
    fake.prompt = undefined;
    expect(await collect(a.sendTurn(input(id)))).toContainEqual({ type: "turn_finished", requestId: "r", reason: "completed" });
  });
  it("closes a timed out runtime and tells the session manager instead of reusing it", async () => {
    fake.prompt = async () => { throw new Error("session/prompt: timed out"); };
    const onRuntimeClosed = vi.fn(), a = adapter(), id = await a.startSession(ctx({ onRuntimeClosed }));
    const events = await collect(a.sendTurn(input(id)));
    expect(events.filter(e => e.type === "turn_finished")).toEqual([{ type: "turn_finished", requestId: "r", reason: "error" }]);
    expect(onRuntimeClosed).toHaveBeenCalledOnce();
    expect(fake.peers[0].close).toHaveBeenCalledOnce();
    expect(await collect(a.sendTurn(input(id)))).toContainEqual(expect.objectContaining({ type: "error", code: "no_session" }));
  });
});

describe("Antigravity tool schemas", () => {
  it("asks for full tool schemas because its model sends no arguments to an open schema", () => {
    expect(AntigravityAdapter.capabilities.fullSchemasOnly).toBe(true);
  });
});
