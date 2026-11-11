import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { ProviderEvent, UserInputResponse } from "../protocol/src/index.js";
import type { StartSessionContext } from "../providers/src/types.js";
import type { McpToolBridge } from "../providers/src/mcp-tool-bridge.js";

const sdk = vi.hoisted(() => ({ config: null as any, client: null as any, session: null as any, handler: null as any,
  run: null as any, open: null as any, resumeId: undefined as string | undefined }));
vi.mock("@github/copilot-sdk", () => ({
  RuntimeConnection: { forStdio: (options: any) => ({ kind: "stdio", ...options }) },
  ToolSet: class {
    items: string[] = [];
    addBuiltIn(value: string | string[]) { this.items.push(...(Array.isArray(value) ? value : [value]).map(x => `builtin:${x}`)); return this; }
    addMcp(value: string) { this.items.push(`mcp:${value}`); return this; }
    addCustom(value: string) { this.items.push(`custom:${value}`); return this; }
    toArray() { return this.items; }
  },
  convertMcpCallToolResult: (result: any) => ({ textResultForLlm: result.content.map((part: any) => part.text).join(""), resultType: result.isError ? "failure" : "success" }),
  CopilotClient: class {
    constructor(options: any) { sdk.client = options; }
    async start() {}
    async createSession(config: any) { sdk.config = config; return sdk.open(config); }
    async resumeSession(id: string, config: any) { sdk.resumeId = id; sdk.config = config; return sdk.open(config); }
    async stop() { sdk.client.stopped = true; }
  },
}));
import { CopilotNativeSessions, resolveCopilotPath } from "../providers/src/copilot-native.js";

const schema = { type: "object", properties: { address: { type: "string" } }, required: ["address"] };
const bridge = { calls: [] as Array<{ name: string; args: unknown; signal?: AbortSignal }>, closed: 0, reply: null as any, fail: null as Error | null };
const fakeBridge: McpToolBridge = {
  // The MCP server lists more than one session may offer; the session keeps only the tools it is allowed to.
  tools: ["read_range", "get_sheets", "get_selection", "read_probe", "not_exposed"].map(name => ({ name, description: `${name} description`, inputSchema: schema })),
  call: async (name, args, signal) => {
    bridge.calls.push({ name, args, signal });
    if (bridge.fail) throw bridge.fail;
    return bridge.reply ?? { content: [{ type: "text", text: `${name} ok` }] };
  },
  close: async () => { bridge.closed++; },
};
let manager: CopilotNativeSessions;
let dir: string;
let ctx: StartSessionContext;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), "copilot-native-test-"));
  vi.stubEnv("EXCEL_AI_HOME", dir);
  vi.stubEnv("COPILOT_HOME", dir);
  manager = new CopilotNativeSessions(async () => fakeBridge);
  Object.assign(bridge, { calls: [], closed: 0, reply: null, fail: null });
  ctx = { workbookConnectionId: "workbook", sessionId: "bridge", signal: new AbortController().signal,
    mcp: { command: "node", args: ["excel.js"], env: {}, bridgeUrl: "", sessionId: "", token: "" },
    requestUserInput: vi.fn(async () => ({ cancelled: false, answers: { answer: ["New sheet"] } })),
  };
  sdk.resumeId = undefined;
  sdk.open = async () => sdk.session;
  sdk.session = { sessionId: "native-id", on: (handler: any) => { sdk.handler = handler; return () => { sdk.handler = null; }; },
    sendAndWait: (...args: any[]) => sdk.run(...args), abort: vi.fn(async () => {}), disconnect: vi.fn(async () => {}) };
  sdk.run = async () => {};
});
afterEach(async () => { await manager.dispose(); vi.unstubAllEnvs(); fs.rmSync(dir, { recursive: true, force: true }); });

async function turn(id: string, signal = new AbortController().signal) {
  const events: ProviderEvent[] = [];
  for await (const event of manager.send({ providerSessionId: id, requestId: "r", prompt: "Ask where", selection: null, signal }, process.env)) events.push(event);
  return events;
}

describe("Copilot native question transport", () => {
  it("uses the installed executable, limits tools, and keeps workbook permission in the gateway", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    expect(manager.has(id)).toBe(true);
    expect(sdk.client.connection.path).toBe(fs.realpathSync(process.execPath));
    expect(sdk.client.connection.args).toContain("--disable-builtin-mcps");
    expect(sdk.config.mcpServers).toBeUndefined();
    expect(sdk.config.toolSearch).toEqual({ enabled: false });
    const tools = sdk.config.availableTools.toArray();
    expect(tools).toContain("builtin:ask_user");
    expect(tools).toContain("custom:read_range");
    expect(tools).not.toContain("builtin:shell");
    expect(sdk.config.onPermissionRequest({ kind: "shell" })).toEqual({ kind: "reject" });
    expect(sdk.config.onPermissionRequest({ kind: "url" })).toEqual({ kind: "reject" });
  });
  it("registers the tools as the host's own tools, not as an MCP server the CLI can block or hide", async () => {
    await manager.start(ctx, process.execPath, process.env);
    expect(sdk.config.mcpServers).toBeUndefined();
    expect(sdk.config.toolSearch).toEqual({ enabled: false });
    // Only what this session may offer, with the schema and description the MCP server gave each tool.
    expect(sdk.config.tools.map((tool: any) => tool.name)).toEqual(["read_range", "get_sheets", "get_selection"]);
    expect(sdk.config.tools[0]).toMatchObject({ description: "read_range description", parameters: schema, skipPermission: true, defer: "never" });
    expect(sdk.config.availableTools.toArray()).toEqual(["builtin:ask_user", "custom:read_range", "custom:get_sheets", "custom:get_selection"]);
  });
  it("offers a client-supplied catalog as listed by the server, and lists it again when a session resumes", async () => {
    ctx.mcp.toolNames = ["read_probe", "get_sheets"];
    await manager.start(ctx, process.execPath, process.env);
    expect(sdk.config.tools.map((tool: any) => tool.name)).toEqual(["get_sheets", "read_probe"]);
    fakeBridge.tools.push({ name: "ext_demo__echo", description: "echo", inputSchema: schema });
    ctx.mcp.toolNames = ["read_probe", "ext_demo__echo"];
    ctx.resumeSessionId = "native-id";
    await manager.start(ctx, process.execPath, process.env);
    expect(sdk.resumeId).toBe("native-id");
    expect(sdk.config.tools.map((tool: any) => tool.name)).toEqual(["read_probe", "ext_demo__echo"]);
    fakeBridge.tools.pop();
  });
  it("runs a host tool through the MCP server and reports failures to the model instead of throwing", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    const [readRange] = sdk.config.tools;
    expect(await readRange.handler({ address: "A1" })).toMatchObject({ textResultForLlm: "read_range ok", resultType: "success" });
    expect(bridge.calls.at(-1)).toMatchObject({ name: "read_range", args: { address: "A1" } });
    bridge.reply = { content: [{ type: "text", text: "no such sheet" }], isError: true };
    expect(await readRange.handler({ address: "A1" })).toMatchObject({ textResultForLlm: "no such sheet", resultType: "failure" });
    bridge.fail = new Error("MCP error -32001: Request timed out");
    expect(await readRange.handler({ address: "A1" })).toEqual({ textResultForLlm: "MCP error -32001: Request timed out", resultType: "failure", error: "MCP error -32001: Request timed out" });
    expect(manager.has(id)).toBe(true);
  });
  it("lets Stop reach a tool call that is waiting on the application", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    const stop = new AbortController();
    sdk.run = async () => { await sdk.config.tools[0].handler({ address: "A1" }); };
    await turn(id, stop.signal);
    expect(bridge.calls.at(-1)?.signal).toBe(stop.signal);
  });
  it("closes the MCP connection with the session", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    expect(bridge.closed).toBe(0);
    await manager.close(id);
    expect(bridge.closed).toBe(1);
  });
  it("returns native answers and streams messages without duplicating final text", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    sdk.run = async () => {
      const answer = await sdk.config.onUserInputRequest({ question: "Where?", choices: ["Current sheet", "New sheet"], allowFreeform: true });
      expect(answer).toEqual({ answer: "New sheet", wasFreeform: false });
      sdk.handler({ type: "assistant.message_delta", data: { messageId: "m", deltaContent: "Done" } });
      sdk.handler({ type: "assistant.message", data: { messageId: "m", content: "Done" } });
    };
    const events = await turn(id);
    expect(ctx.requestUserInput).toHaveBeenCalledWith(expect.objectContaining({ questions: [expect.objectContaining({ id: "answer", question: "Where?" })] }), expect.any(AbortSignal));
    expect(events.filter(e => e.type === "message_delta")).toHaveLength(1);
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "completed" });
  });
  it("does not repeat text that streams under one message id and completes under another", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    sdk.run = async () => {
      for (const chunk of ["Bước 0 — ", "Tool và model"]) sdk.handler({ type: "assistant.message_delta", data: { messageId: "reasoning", deltaContent: chunk } });
      sdk.handler({ type: "assistant.message", data: { messageId: "reasoning", content: "", toolRequests: [] } });
      sdk.handler({ type: "assistant.message", data: { messageId: "final", content: "Bước 0 — Tool và model" } });
    };
    const events = await turn(id);
    const shown = events.filter(e => e.type === "message_delta").map(e => (e as { text: string }).text).join("");
    expect(shown).toBe("Bước 0 — Tool và model");
    expect(events.filter(e => e.type === "message_completed")).toEqual([{ type: "message_completed", requestId: "r", text: "Bước 0 — Tool và model" }]);
  });
  it("appends only the unstreamed remainder of a final message", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    sdk.run = async () => {
      sdk.handler({ type: "assistant.message_delta", data: { messageId: "a", deltaContent: "Hello " } });
      sdk.handler({ type: "assistant.message", data: { messageId: "b", content: "Hello world" } });
    };
    const shown = (await turn(id)).filter(e => e.type === "message_delta").map(e => (e as { text: string }).text).join("");
    expect(shown).toBe("Hello world");
  });
  it("streams reasoning as thinking, one paragraph per block, and shows an unstreamed block whole", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    sdk.run = async () => {
      for (const chunk of ["Check ", "the sheet."]) sdk.handler({ type: "assistant.reasoning_delta", data: { reasoningId: "a", deltaContent: chunk } });
      sdk.handler({ type: "assistant.reasoning", data: { reasoningId: "a", content: "Check the sheet." } });
      sdk.handler({ type: "assistant.reasoning", data: { reasoningId: "b", content: "Then sort." } });
      sdk.handler({ type: "assistant.message", data: { messageId: "x", content: "Done" } });
    };
    const events = await turn(id);
    expect(events.flatMap(e => e.type === "thinking_delta" ? [e.text] : [])).toEqual(["Check ", "the sheet.", "\n\nThen sort."]);
    expect(events.findIndex(e => e.type === "thinking_delta")).toBeLessThan(events.findIndex(e => e.type === "message_delta"));
  });
  it("shows a final message that was never streamed", async () => {
    const id = await manager.start(ctx, process.execPath, process.env);
    sdk.run = async () => { sdk.handler({ type: "assistant.message", data: { messageId: "x", content: "Short answer" } }); };
    const events = await turn(id);
    expect(events.filter(e => e.type === "message_delta")).toEqual([{ type: "message_delta", requestId: "r", text: "Short answer" }]);
  });
  it("rejects skipped native questions without returning a fabricated answer", async () => {
    ctx.requestUserInput = async (): Promise<UserInputResponse> => ({ cancelled: true, answers: {} });
    const id = await manager.start(ctx, process.execPath, process.env);
    sdk.run = async () => {
      await expect(sdk.config.onUserInputRequest({ question: "Where?" })).rejects.toThrow("cancelled");
    };
    expect((await turn(id)).at(-1)).toMatchObject({ reason: "completed" });
  });
  it("resumes the same native session with the question handler registered", async () => {
    ctx.resumeSessionId = "native-id";
    await manager.start(ctx, process.execPath, process.env);
    expect(sdk.resumeId).toBe("native-id");
    expect(sdk.config.onUserInputRequest).toBeTypeOf("function");
  });
  it.each([[false, false], [false, true], [true, false], [true, true]])("omits unsupported reasoning effort for Auto (API=%s, resume=%s)", async (api, resume) => {
    if (api) ctx.mcp.toolNames = ["read_probe"];
    ctx.model = null;
    ctx.effort = "high";
    if (resume) ctx.resumeSessionId = "native-id";
    await manager.start(ctx, process.execPath, process.env);
    expect(sdk.config).not.toHaveProperty("reasoningEffort");
    expect(sdk.resumeId).toBe(resume ? "native-id" : undefined);
  });
  it.each([false, true])("preserves explicit API model and add-in effort settings (%s)", async api => {
    ctx.effort = "high";
    ctx.model = "gpt-6-luna";
    if (api) ctx.mcp.toolNames = ["read_probe"];
    await manager.start(ctx, process.execPath, process.env);
    expect(sdk.config.reasoningEffort).toBe("high");
  });
  it.each([false, true])("retries unsupported model effort without changing model or conversation (resume=%s)", async resume => {
    ctx.model = "gpt-4.1"; ctx.effort = "low";
    if (resume) ctx.resumeSessionId = "native-id";
    const configs: any[] = [];
    sdk.open = async (config: any) => {
      configs.push({ ...config });
      if (config.reasoningEffort) throw new Error("Model 'gpt-4.1' does not support reasoning effort configuration.");
      return sdk.session;
    };
    await manager.start(ctx, process.execPath, process.env);
    expect(configs).toHaveLength(2);
    expect(configs[0]).toMatchObject({ model: "gpt-4.1", reasoningEffort: "low" });
    expect(configs[1].model).toBe("gpt-4.1");
    expect(configs[1]).not.toHaveProperty("reasoningEffort");
    expect(sdk.resumeId).toBe(resume ? "native-id" : undefined);
  });
  it("does not retry authentication failures as unsupported effort", async () => {
    ctx.model = "gpt-4.1"; ctx.effort = "low";
    sdk.open = vi.fn(async () => { throw new Error("Authentication failed"); });
    await expect(manager.start(ctx, process.execPath, process.env)).rejects.toThrow("Authentication failed");
    expect(sdk.open).toHaveBeenCalledTimes(1);
    expect(sdk.client.stopped).toBe(true);
    expect(bridge.closed).toBe(1);
  });
  it.each([["word", false], ["powerpoint", false], ["excel", true]])("exposes only %s tools through the host fallback (Excel tools allowed: %s)", async (host, excel) => {
    ctx.mcp.env = { OFFICE_HOST: host };
    await manager.start(ctx, process.execPath, process.env);
    const tools = sdk.config.availableTools.toArray();
    expect(tools).toContain("builtin:ask_user");
    expect(tools).toContain("custom:get_selection");
    expect(tools.includes("custom:read_range")).toBe(excel);
  });
  it("refuses to start for a host it does not know", async () => {
    ctx.mcp.env = { OFFICE_HOST: "outlook" };
    await expect(manager.start(ctx, process.execPath, process.env)).rejects.toThrow(/OFFICE_HOST/);
  });
  it("rejects requests outside a running turn and resolves PATH commands", async () => {
    await manager.start(ctx, process.execPath, process.env);
    await expect(sdk.config.onUserInputRequest({ question: "Stale?" })).rejects.toThrow("cancelled");
    expect(resolveCopilotPath(path.basename(process.execPath), { PATH: path.dirname(process.execPath) })).toBe(fs.realpathSync(process.execPath));
  });
  it("invokes an informational native command without sending it to the model", async () => {
    const invoke = vi.fn(async () => ({ kind: "text", text: "Available skills" }));
    sdk.session.rpc = { commands: { invoke, list: async () => ({ commands: [{ name: "skills", description: "Skills" }] }) } };
    sdk.run = vi.fn();
    const id = await manager.start(ctx, process.execPath, process.env);
    expect(await manager.listCommands(process.execPath, process.env, id)).toEqual([expect.objectContaining({ name: "skills", execution: "native" })]);
    const events: ProviderEvent[] = [];
    for await (const event of manager.send({ providerSessionId: id, requestId: "cmd", prompt: "INJECTED CONTEXT", selection: null,
      command: { name: "skills", args: "list" }, signal: new AbortController().signal }, process.env)) events.push(event);
    expect(invoke).toHaveBeenCalledWith({ name: "skills", input: "list" });
    expect(sdk.run).not.toHaveBeenCalled();
    expect(events).toContainEqual({ type: "message_completed", requestId: "cmd", text: "Available skills" });
    expect(events.at(-1)).toMatchObject({ type: "turn_finished", reason: "completed" });
  });
  it("submits the provider's resolved agent prompt for action commands", async () => {
    sdk.session.rpc = { commands: { invoke: vi.fn(async () => ({ kind: "agent-prompt", prompt: "Review the changes", mode: "interactive" })) } };
    sdk.run = vi.fn(async () => {});
    const id = await manager.start(ctx, process.execPath, process.env);
    for await (const _ of manager.send({ providerSessionId: id, requestId: "cmd", prompt: "/review", selection: null,
      command: { name: "review", args: "" }, signal: new AbortController().signal }, process.env)) { /* drain */ }
    expect(sdk.run).toHaveBeenCalledWith({ prompt: "Review the changes", agentMode: "interactive" }, expect.any(Number));
  });
  it("R10 resolves copied Windows Path and PathExt variables case-insensitively", () => {
    const executable = path.join(dir, "copilot.TEST"); fs.writeFileSync(executable, "");
    vi.stubGlobal("process", { ...process, platform: "win32" });
    try { expect(resolveCopilotPath("copilot", { Path: dir, PathExt: ".TEST" })).toBe(fs.realpathSync(executable)); }
    finally { vi.unstubAllGlobals(); }
  });
  it.skipIf(process.platform !== "win32")("resolves Windows PATH commands with or without an extension", () => {
    const executable = path.join(dir, "copilot.EXE");
    fs.writeFileSync(executable, "");
    fs.writeFileSync(path.join(dir, "copilot"), "#!/bin/sh\n");
    fs.writeFileSync(`${executable}.EXE`, "");
    const env = { PATH: dir, PATHEXT: ".EXE;.CMD;.BAT" };
    expect(resolveCopilotPath("copilot.EXE", env)).toBe(fs.realpathSync(executable));
    expect(resolveCopilotPath("copilot", env)).toBe(fs.realpathSync(executable));
  });
});
