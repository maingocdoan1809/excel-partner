import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import { createInterface } from "node:readline";
import type { ProviderEvent } from "../protocol/src/index.js";
import type { StartSessionContext } from "../providers/src/types.js";

const cli = vi.hoisted(() => ({ spawn: vi.fn(), messages: [] as object[], failInit: false, hold: false,
  inputs: [] as any[], children: [] as any[] }));
vi.mock("cross-spawn", () => ({ default: cli.spawn }));
import { ClaudeAdapter } from "../providers/src/claude.js";

const runtime = { bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } };
let adapter: ClaudeAdapter;
beforeEach(() => {
  adapter = new ClaudeAdapter(runtime);
  vi.spyOn(adapter, "sessionTitle").mockResolvedValue(undefined);
  cli.messages = [{ type: "result", session_id: "confirmed-session", is_error: false }];
  cli.failInit = false;
  cli.hold = false;
  cli.inputs = [];
  cli.children = [];
  cli.spawn.mockReset().mockImplementation(() => {
    const child = Object.assign(new EventEmitter(), {
      stdin: new PassThrough(), stdout: new PassThrough(), stderr: new PassThrough(),
      exitCode: null as number | null, signalCode: null as string | null,
      kill: vi.fn((signal: string) => {
        child.signalCode = signal;
        child.stdout.end(); child.stderr.end();
        child.emit("exit", null, signal);
        queueMicrotask(() => child.emit("close", null, signal));
        return true;
      }),
    });
    cli.children.push(child);
    createInterface({ input: child.stdin }).on("line", line => {
      const message = JSON.parse(line);
      cli.inputs.push(message);
      setImmediate(() => {
        if (child.signalCode) return;
        if (message.type === "control_request") {
          child.stdout.write(JSON.stringify({ type: "control_response", response: {
            request_id: message.request_id, subtype: cli.failInit ? "error" : "success",
            ...(cli.failInit ? { error: "initialization failed" } : { response: {} }),
          } }) + "\n");
          if (message.request.subtype === "interrupt") child.stdout.write(JSON.stringify({ type: "result", is_error: true }) + "\n");
        } else if (message.type === "user" && !cli.hold) {
          for (const response of cli.messages) child.stdout.write(JSON.stringify(response) + "\n");
        }
      });
    });
    return child;
  });
});
afterEach(async () => { await adapter.dispose(); vi.restoreAllMocks(); });

async function start(resumeSessionId?: string, overrides: Partial<StartSessionContext> = {}) {
  return adapter.startSession({ workbookConnectionId: "workbook", sessionId: "bridge", resumeSessionId,
    signal: new AbortController().signal,
    mcp: { command: process.execPath, args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" }, ...overrides,
  });
}
async function turn(providerSessionId: string, signal = new AbortController().signal, requestId = "request") {
  const events: ProviderEvent[] = [];
  for await (const event of adapter.sendTurn({ providerSessionId, requestId, prompt: "hello", selection: null, signal })) events.push(event);
  return events;
}

describe("Claude persistent sessions", () => {
  it.each([false, true])("exposes native web tools only when Web search is enabled (%j)", async enabled => {
    await start(undefined, { mcp: { command: process.execPath, args: [], env: enabled ? { EXCEL_WEB_SEARCH: "1" } : {}, bridgeUrl: "", sessionId: "", token: "" } });
    const args = cli.spawn.mock.calls[0]![1] as string[];
    expect(args[args.indexOf("--tools") + 1]).toBe(enabled ? "AskUserQuestion,WebSearch,WebFetch" : "AskUserQuestion");
    expect(args.includes("WebSearch")).toBe(enabled);
    expect(args.includes("WebFetch")).toBe(enabled);
  });

  it.each([undefined, "", " ", "Custom instructions"])("only overrides native instructions for a non-empty add-in prompt (%j)", async systemPrompt => {
    await start(undefined, { systemPrompt });
    const args = cli.spawn.mock.calls[0]![1] as string[];
    if (systemPrompt?.trim()) expect(args[args.indexOf("--system-prompt") + 1]).toBe(systemPrompt);
    else expect(args).not.toContain("--system-prompt");
  });

  it("does not launch when session creation was already cancelled", async () => {
    const controller = new AbortController(); controller.abort();
    await expect(start(undefined, { signal: controller.signal })).rejects.toThrow();
    expect(cli.spawn).not.toHaveBeenCalled();
  });

  it("does not publish a locally generated ID for a cancelled first turn", async () => {
    const id = await start();
    const controller = new AbortController(); controller.abort();
    const events = await turn(id, controller.signal);
    expect(events.some(event => event.type === "session_info")).toBe(false);
    expect(events.at(-1)).toMatchObject({ reason: "cancelled" });
    expect(cli.inputs.filter(m => m.type === "user")).toHaveLength(0);
  });

  it("closes a process whose initialization fails", async () => {
    cli.failInit = true;
    await expect(start()).rejects.toThrow("initialization failed");
    expect(cli.children[0].kill).toHaveBeenCalled();
  });

  it("reuses one process and keeps request IDs and usage separate across turns", async () => {
    cli.messages = [{ type: "result", session_id: "confirmed-session", usage: { input_tokens: 10, output_tokens: 5 } }];
    const id = await start();
    const first = await turn(id, undefined, "first");
    const second = await turn(id, undefined, "second");
    expect(cli.spawn).toHaveBeenCalledTimes(1);
    expect(cli.inputs.filter(m => m.type === "user")).toHaveLength(2);
    expect(cli.children[0].stdin.writableEnded).toBe(false);
    expect(cli.children[0].kill).not.toHaveBeenCalled();
    for (const [events, requestId] of [[first, "first"], [second, "second"]] as const) {
      expect(events.every(e => e.requestId === requestId)).toBe(true);
      expect(events.filter(e => e.type === "turn_finished")).toHaveLength(1);
      expect(events.find(e => e.type === "usage_update")).toMatchObject({ usage: { totalTokens: 15 } });
      expect(events.find(e => e.type === "session_info")).toMatchObject({ providerSessionId: "confirmed-session" });
    }
  });

  it.each(["init", "result"])("publishes only a CLI-confirmed ID from %s", async kind => {
    cli.messages = [...(kind === "init" ? [{ type: "system", subtype: "init", session_id: "confirmed-session" }] : []),
      { type: "result", ...(kind === "result" ? { session_id: "confirmed-session" } : {}) }];
    expect((await turn(await start())).find(e => e.type === "session_info")).toMatchObject({ providerSessionId: "confirmed-session" });
  });

  it("uses --resume only when opening a saved conversation", async () => {
    const id = await start("confirmed-session");
    await turn(id); await turn(id);
    expect(cli.spawn).toHaveBeenCalledTimes(1);
    const args = cli.spawn.mock.calls[0]![1] as string[];
    expect(args[args.indexOf("--resume") + 1]).toBe("confirmed-session");
    expect(args).not.toContain("--session-id");
  });

  it("rejects a mismatched resumed identity", async () => {
    const events = await turn(await start("different-session"));
    expect(events.find(e => e.type === "error")).toMatchObject({ message: expect.stringContaining("different session") });
    expect(events.at(-1)).toMatchObject({ reason: "error" });
  });

  it("interrupts a turn and accepts the next prompt on the same process", async () => {
    const id = await start();
    cli.hold = true;
    const controller = new AbortController();
    const pending = turn(id, controller.signal);
    await vi.waitFor(() => expect(cli.inputs.some(m => m.type === "user")).toBe(true));
    controller.abort();
    expect((await pending).at(-1)).toMatchObject({ reason: "cancelled" });
    cli.hold = false;
    expect((await turn(id)).at(-1)).toMatchObject({ reason: "completed" });
    expect(cli.spawn).toHaveBeenCalledTimes(1);
    expect(cli.inputs.filter(m => m.request?.subtype === "interrupt")).toHaveLength(1);
  });

  it("completes a crashed active turn with an error", async () => {
    const id = await start(); cli.hold = true;
    const pending = turn(id);
    await vi.waitFor(() => expect(cli.inputs.some(m => m.type === "user")).toBe(true));
    cli.children[0].kill("SIGTERM");
    const events = await pending;
    expect(events.filter(e => e.type === "turn_finished")).toHaveLength(1);
    expect(events.at(-1)).toMatchObject({ reason: "error" });
  });
});
