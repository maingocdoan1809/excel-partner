import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import { createInterface } from "node:readline";
import { ClaudeAdapter } from "../providers/src/claude.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { ProviderSessionManager } from "../providers/src/session-manager.js";
import { WorkbookSession } from "../bridge/src/session.js";
import type { ProviderEvent } from "../protocol/src/index.js";

/** A message the person sends during a turn goes to the provider itself (Claude: written to the CLI's input). */

const cli = vi.hoisted(() => ({ spawn: vi.fn(), children: [] as any[], written: [] as any[] }));
vi.mock("cross-spawn", () => ({ default: cli.spawn }));

let adapter: ClaudeAdapter;
beforeEach(() => {
  adapter = new ClaudeAdapter({ bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });
  vi.spyOn(adapter, "sessionTitle").mockResolvedValue(undefined);
  cli.children = []; cli.written = [];
  cli.spawn.mockReset().mockImplementation(() => {
    const child = Object.assign(new EventEmitter(), {
      stdin: new PassThrough(), stdout: new PassThrough(), stderr: new PassThrough(),
      exitCode: null as number | null, signalCode: null as string | null,
      kill: vi.fn((signal: string) => { child.signalCode = signal; child.emit("exit", null, signal); child.stdout.end(); child.stderr.end(); queueMicrotask(() => child.emit("close", null, signal)); return true; }),
    });
    cli.children.push(child);
    // The fake CLI never answers by itself: each test plays the CLI's side by hand.
    createInterface({ input: child.stdin }).on("line", line => {
      const message = JSON.parse(line);
      cli.written.push(message);
      if (message.type === "control_request") {
        queueMicrotask(() => child.stdout.write(JSON.stringify({ type: "control_response", response: { request_id: message.request_id, subtype: "success", response: {} } }) + "\n"));
      }
    });
    return child;
  });
});
afterEach(async () => { await adapter.dispose(); vi.useRealTimers(); vi.restoreAllMocks(); });

const tick = () => new Promise(resolve => setImmediate(resolve));
const emit = (message: object) => cli.children[0].stdout.write(JSON.stringify(message) + "\n");
const result = () => ({ type: "result", session_id: "native", is_error: false });
const answer = (text: string) => ({ type: "assistant", message: { id: `m-${text}`, content: [{ type: "text", text }] } });

async function startTurn() {
  const manager = new ProviderSessionManager(new ProviderRegistry().register(adapter), async () => {}, { debug() {}, warn() {} });
  const session = new WorkbookSession("workbook"); session.provider = "claude";
  const settings = { model: null, effort: "medium" as const, capabilities: { webSearch: false, playwright: false }, systemPromptRevision: "r" };
  const context = () => ({ workbookConnectionId: session.workbookConnectionId, sessionId: session.id,
    signal: new AbortController().signal, mcp: { command: "node", args: [], env: {}, bridgeUrl: "", sessionId: "", token: "" } });
  await manager.ensure(session, settings, context);
  const events: ProviderEvent[] = [];
  let finished = false;
  const running = (async () => {
    for await (const event of manager.sendTurn(session, { requestId: "r-1", prompt: "hello", selection: null, signal: new AbortController().signal })) events.push(event);
    finished = true;
  })();
  await tick();
  return { session, events, running, isFinished: () => finished, steer: (text: string) => manager.steer(session, text) };
}

const steeredMessage = () => cli.written.find(message => message.type === "user" && message.message.content === "Also use blue");

it("writes a message sent during the turn to the CLI with an id, and the turn ends once, with its answer", async () => {
  const turn = await startTurn();
  expect(await turn.steer("Also use blue")).toBe(true);
  expect(steeredMessage()).toMatchObject({ type: "user", uuid: expect.any(String), parent_tool_use_id: null });
  // The CLI takes it at its next step (it says so by repeating it), then answers inside the same turn.
  emit({ type: "user", isReplay: true, uuid: steeredMessage().uuid, message: { role: "user", content: "Also use blue" } });
  emit(answer("Blue it is"));
  emit(result());
  await turn.running;
  expect(turn.events.filter(event => event.type === "turn_finished")).toHaveLength(1);
  expect(turn.events).toContainEqual(expect.objectContaining({ type: "message_completed", text: "Blue it is" }));
});

it("keeps the turn open when the CLI only takes the message after its last result, and ends it after the answer", async () => {
  const turn = await startTurn();
  await turn.steer("Also use blue");
  emit(answer("ONE"));
  emit(result()); // the CLI had no step left: the message runs as a turn of its own
  await tick(); await tick();
  expect(turn.isFinished()).toBe(false);
  emit({ type: "user", isReplay: true, uuid: steeredMessage().uuid, message: { role: "user", content: "Also use blue" } });
  emit(answer("TWO"));
  emit(result());
  await turn.running;
  expect(turn.events.filter(event => event.type === "turn_finished")).toHaveLength(1);
  expect(turn.events.at(-1)).toMatchObject({ type: "turn_finished", reason: "completed" });
  expect(turn.events.filter(event => event.type === "message_completed").map(event => (event as { text: string }).text)).toEqual(["ONE", "TWO"]);
});

it("does not wait forever for a message the CLI never takes", async () => {
  vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
  const turn = await startTurn();
  await turn.steer("Also use blue");
  emit(result());
  await vi.advanceTimersByTimeAsync(100);
  expect(turn.isFinished()).toBe(false);
  await vi.advanceTimersByTimeAsync(15_000);
  await turn.running;
  expect(turn.events.at(-1)).toMatchObject({ type: "turn_finished", reason: "completed" });
});

it("says no when there is no turn to join, so the caller queues the message", async () => {
  const turn = await startTurn();
  emit(result());
  await turn.running;
  expect(await turn.steer("Too late")).toBe(false);
  expect(cli.written.some(message => message.type === "user" && message.message.content === "Too late")).toBe(false);
});
