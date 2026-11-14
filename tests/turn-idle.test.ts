import { afterEach, describe, expect, it, vi } from "vitest";
import type { ServerMessage } from "../protocol/src/index.js";
import { WorkbookSession } from "../bridge/src/session.js";
import { SessionLimitError, WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import { TurnRunner, turnIdleMs, type TurnRunnerDeps } from "../bridge/src/turns.js";
import { Logger } from "../bridge/src/logger.js";

afterEach(() => { vi.useRealTimers(); vi.unstubAllEnvs(); });

/** A pane whose provider starts a turn and then says nothing until it is cancelled. */
function silentTurn() {
  const sent: ServerMessage[] = [];
  const session = new WorkbookSession("0123456789abcdef", { host: "excel" } as never);
  session.attachSocket({ readyState: 1, send: (text: string) => sent.push(JSON.parse(text)), close() {} } as never);
  session.provider = "claude";
  session.providerSessionId = "runtime";
  const cancel = vi.fn(async () => undefined);
  const providerSessions = {
    ensure: async () => undefined, cancel, close: async () => undefined,
    async *sendTurn(_session: unknown, input: { requestId: string; signal: AbortSignal }) {
      yield { type: "turn_started", requestId: input.requestId };
      await new Promise(resolve => input.signal.addEventListener("abort", resolve, { once: true }));
    },
  };
  const skills = { get: () => undefined, list: () => [], listTemplates: () => [] };
  const deps = {
    config: { cancelGraceMs: 5_000 }, logger: new Logger("test", "error", () => {}),
    providers: { require: () => ({ capabilities: {} }), get: () => ({ capabilities: {} }) },
    providerSessions, delegations: { observe() {}, finish() {}, accept: () => true, decline() {} },
    externalTools: { closeSession: async () => undefined }, usage: { append() {} }, skills,
    systemPrompt: { forHost: () => ({ read: () => ({ content: "", revision: "r" }) }) },
    savedReferences: { prompt: () => "" }, references: {}, discoverMcp: async () => new Map(), hasOtherFiles: () => false,
  } as unknown as TurnRunnerDeps;
  const runner = new TurnRunner(deps);
  const run = runner.run(session, { requestId: "r1", prompt: "hello", selection: null });
  return { session, sent, run, cancel };
}

describe("turn idle watchdog", () => {
  it("reads the idle limit from OFFICE_PARTNER_TURN_IDLE_MINUTES, 20 minutes by default", () => {
    expect(turnIdleMs({})).toBe(20 * 60_000);
    expect(turnIdleMs({ OFFICE_PARTNER_TURN_IDLE_MINUTES: "5" })).toBe(5 * 60_000);
    expect(turnIdleMs({ OFFICE_PARTNER_TURN_IDLE_MINUTES: "nonsense" })).toBe(20 * 60_000);
  });

  it("stops a turn that shows no activity with TURN_IDLE_TIMEOUT", async () => {
    vi.useFakeTimers();
    vi.stubEnv("OFFICE_PARTNER_TURN_IDLE_MINUTES", "2");
    const w = silentTurn();
    await vi.advanceTimersByTimeAsync(60_000);
    expect(w.sent.some(m => m.type === "error")).toBe(false);
    await vi.advanceTimersByTimeAsync(60_000);
    await w.run;
    expect(w.sent).toContainEqual(expect.objectContaining({ type: "error", code: "TURN_IDLE_TIMEOUT" }));
    expect(w.sent).toContainEqual({ type: "turn_finished", requestId: "r1", reason: "error" });
    expect(w.sent.filter(m => m.type === "error")).toHaveLength(1);
    expect(w.cancel).toHaveBeenCalled();
    expect(w.session.isBusy).toBe(false);
  });

  it("does not count waiting for the user's answer as idle", async () => {
    vi.useFakeTimers();
    vi.stubEnv("OFFICE_PARTNER_TURN_IDLE_MINUTES", "1");
    const w = silentTurn();
    await vi.advanceTimersByTimeAsync(0);
    const answer = w.session.requestUserInput({ questions: [{ id: "q", question: "Which sheet?", options: [{ label: "A" }, { label: "B" }] }] } as never);
    await vi.advanceTimersByTimeAsync(5 * 60_000);
    expect(w.sent.some(m => m.type === "error")).toBe(false);
    const request = w.sent.find(m => m.type === "tool_request") as { callId: string };
    w.session.resolveTool(request.callId, { ok: true, data: { cancelled: false, answers: { q: ["A"] } } });
    await answer;
    await vi.advanceTimersByTimeAsync(2 * 60_000);
    await w.run;
    expect(w.sent).toContainEqual(expect.objectContaining({ type: "error", code: "TURN_IDLE_TIMEOUT" }));
  });
});

describe("session limit", () => {
  it("evicts the least recently active idle session, never a busy one, and refuses when all are busy", () => {
    const sessions = new WorkbookSessions({ ttlMs: 60_000, maxSessions: 2 });
    const busy = sessions.create("aaaaaaaaaaaaaaaa", null);
    busy.beginTurn("r1");
    busy.lastActivityAt = 0;
    const idle = sessions.create("bbbbbbbbbbbbbbbb", null);
    sessions.create("cccccccccccccccc", null);
    expect(sessions.get(busy.id)).toBe(busy);
    expect(sessions.get(idle.id)).toBeUndefined();
    sessions.list().forEach(session => session.beginTurn("r"));
    expect(() => sessions.create("dddddddddddddddd", null)).toThrow(SessionLimitError);
    expect(sessions.size).toBe(2);
  });
});
