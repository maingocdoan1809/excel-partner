import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import { createInterface } from "node:readline";
import { ClaudeAdapter } from "../providers/src/claude.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { ProviderSessionManager } from "../bridge/src/provider-session-manager.js";
import { WorkbookSession } from "../bridge/src/session.js";
import type { ProviderEvent } from "../protocol/src/index.js";
import { ClaudeProcess } from "../providers/src/claude-process.js";

const cli = vi.hoisted(() => ({ spawn: vi.fn(), hold: false, interrupt: "reject", children: [] as any[] }));
vi.mock("cross-spawn", () => ({ default: cli.spawn }));
let adapter: ClaudeAdapter;
beforeEach(() => {
  adapter = new ClaudeAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
  vi.spyOn(adapter, "sessionTitle").mockResolvedValue(undefined);
  cli.hold = false; cli.interrupt = "reject"; cli.children = [];
  cli.spawn.mockReset().mockImplementation(() => {
    const child = Object.assign(new EventEmitter(), {
      stdin: new PassThrough(), stdout: new PassThrough(), stderr: new PassThrough(),
      exitCode: null as number | null, signalCode: null as string | null,
      kill: vi.fn((signal: string) => {
        child.signalCode = signal;
        child.emit("exit", null, signal);
        child.stdout.end(); child.stderr.end();
        queueMicrotask(() => child.emit("close", null, signal));
        return true;
      }),
    });
    cli.children.push(child);
    const emit = (message: object) => child.stdout.write(JSON.stringify(message) + "\n");
    createInterface({ input: child.stdin }).on("line", line => {
      const message = JSON.parse(line);
      queueMicrotask(() => {
        if (message.type === "control_request") {
          if (message.request.subtype === "interrupt" && cli.interrupt === "noack-result") {
            emit({ type: "result", session_id: "native", is_error: true }); return;
          }
          emit({ type: "control_response", response: { request_id: message.request_id,
            subtype: message.request.subtype === "interrupt" ? "error" : "success", error: "interrupt unsupported", response: {} } });
        } else if (message.type === "user" && !cli.hold) emit({ type: "result", session_id: "native", is_error: false });
      });
    });
    return child;
  });
});
afterEach(async () => { await adapter.dispose(); vi.useRealTimers(); vi.restoreAllMocks(); });

async function setup() {
  const manager = new ProviderSessionManager(new ProviderRegistry().register(adapter), async () => {}, { debug() {}, warn() {} });
  const session = new WorkbookSession("workbook"); session.provider = "claude";
  const settings = { model: null, effort: "medium" as const, capabilities: { webSearch: false, playwright: false }, systemPromptRevision: "r" };
  const context = () => ({ workbookConnectionId: session.workbookConnectionId, sessionId: session.id,
    signal: new AbortController().signal, mcp: { command: "node", args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" } });
  await manager.ensure(session, settings, context);
  let seq = 0;
  const turn = async (signal = new AbortController().signal) => {
    const events: ProviderEvent[] = [];
    for await (const e of manager.sendTurn(session, { requestId: `r-${++seq}`, prompt: "hello", selection: null, signal })) events.push(e);
    return events;
  };
  await turn();
  return { manager, session, settings, context, turn };
}

it("reopens the runtime before the next prompt after interrupt fallback kills Claude", async () => {
  const h = await setup(); cli.hold = true;
  const controller = new AbortController();
  const running = h.turn(controller.signal);
  await new Promise(resolve => setImmediate(resolve));
  controller.abort();
  expect((await running).at(-1)).toMatchObject({ reason: "cancelled" });
  expect(cli.children[0].kill).toHaveBeenCalled();
  await h.manager.ensure(h.session, { ...h.settings, resumeSessionId: "native" }, h.context);
  expect(cli.spawn).toHaveBeenCalledTimes(2);
  cli.hold = false;
  expect((await h.turn()).at(-1)).toMatchObject({ reason: "completed" });
});

it("does not let an old interrupt timeout kill the next running turn", async () => {
  const h = await setup(); cli.hold = true; cli.interrupt = "noack-result";
  vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
  const controller = new AbortController();
  const running = h.turn(controller.signal);
  await new Promise(resolve => setImmediate(resolve));
  controller.abort();
  expect((await running).at(-1)).toMatchObject({ reason: "cancelled" });
  expect(vi.getTimerCount()).toBe(0);
  const next = h.turn();
  await new Promise(resolve => setImmediate(resolve));
  await vi.advanceTimersByTimeAsync(5001);
  const killed = cli.children[0].kill.mock.calls.length;
  // Always release the pending iterator, including when the expected behavior works.
  cli.children[0].stdout.write(JSON.stringify({ type: "result", session_id: "native", is_error: false }) + "\n");
  await next;
  expect(killed).toBe(0);
});

it("drains a final stdout result delivered after the process exit event", async () => {
  const received: unknown[] = [];
  const transport = new ClaudeProcess({ command: "claude", args: [], cwd: ".", env: {},
    onMessage: m => received.push(m), onExit() { received.push("closed"); }, onStderr() {} });
  const child = cli.children[0];
  child.exitCode = 0;
  child.emit("exit", 0, null);
  child.stdout.end(JSON.stringify({ type: "result", session_id: "native", is_error: false }) + "\n");
  child.emit("close", 0, null);
  await transport.close();
  expect(received).toEqual([{ type: "result", session_id: "native", is_error: false }, "closed"]);
});
