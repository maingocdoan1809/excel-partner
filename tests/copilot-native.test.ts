import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { ProviderEvent, UserInputResponse } from "../protocol/src/index.js";
import type { StartSessionContext } from "../providers/src/types.js";

const sdk = vi.hoisted(() => ({ config: null as any, client: null as any, session: null as any, handler: null as any,
  run: null as any, resumeId: undefined as string | undefined }));
vi.mock("@github/copilot-sdk", () => ({
  RuntimeConnection: { forStdio: (options: any) => ({ kind: "stdio", ...options }) },
  ToolSet: class {
    items: string[] = [];
    addBuiltIn(value: string | string[]) { this.items.push(...(Array.isArray(value) ? value : [value]).map(x => `builtin:${x}`)); return this; }
    addMcp(value: string) { this.items.push(`mcp:${value}`); return this; }
    toArray() { return this.items; }
  },
  CopilotClient: class {
    constructor(options: any) { sdk.client = options; }
    async start() {}
    async createSession(config: any) { sdk.config = config; return sdk.session; }
    async resumeSession(id: string, config: any) { sdk.resumeId = id; sdk.config = config; return sdk.session; }
    async stop() { sdk.client.stopped = true; }
  },
}));
import { CopilotNativeSessions, resolveCopilotPath } from "../providers/src/copilot-native.js";

let manager: CopilotNativeSessions;
let dir: string;
let ctx: StartSessionContext;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), "copilot-native-test-"));
  vi.stubEnv("EXCEL_AI_HOME", dir);
  vi.stubEnv("COPILOT_HOME", dir);
  manager = new CopilotNativeSessions();
  ctx = { workbookConnectionId: "workbook", sessionId: "bridge", signal: new AbortController().signal,
    mcp: { command: "node", args: ["excel.js"], env: {}, bridgeUrl: "", sessionId: "", token: "" },
    requestUserInput: vi.fn(async () => ({ cancelled: false, answers: { answer: ["New sheet"] } })),
  };
  sdk.resumeId = undefined;
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
    const tools = sdk.config.availableTools.toArray();
    expect(tools).toContain("builtin:ask_user");
    expect(tools).toContain("mcp:excel-read_range");
    expect(tools).not.toContain("builtin:shell");
    expect(sdk.config.onPermissionRequest({ kind: "mcp", serverName: "excel" })).toEqual({ kind: "approve-once" });
    expect(sdk.config.onPermissionRequest({ kind: "shell" })).toEqual({ kind: "reject" });
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
