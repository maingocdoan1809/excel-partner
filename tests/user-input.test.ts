import { describe, expect, it } from "vitest";
import type { WebSocket } from "ws";
import { UserInputRequestSchema, validateUserInputResponse, type ServerMessage } from "../protocol/src/index.js";
import { WorkbookSession } from "../bridge/src/session.js";
import { answerClaudeQuestion, answerCodexQuestions } from "../providers/src/user-input.js";
import { SessionManager } from "../bridge/src/session-manager.js";

const request = UserInputRequestSchema.parse({ questions: [{ id: "output", question: "Where?",
  options: [{ label: "Current sheet" }, { label: "New sheet" }], allowFreeform: false }] });

function connected() {
  const messages: ServerMessage[] = [];
  const manager = new SessionManager({ ttlMs: 60_000 });
  const session: WorkbookSession = manager.create("workbook", null);
  const socket = { readyState: 1, send: (raw: string) => messages.push(JSON.parse(raw)) } as unknown as WebSocket;
  session.attachSocket(socket);
  session.beginTurn("turn");
  return { session, messages, socket };
}

describe("native provider questions", () => {
  it("preserves Codex IDs and returns its native answer envelope", async () => {
    expect(await answerCodexQuestions(request, async received => {
      expect(received.questions[0]?.id).toBe("output");
      expect(received.questions[0]?.allowFreeform).toBe(false);
      return { cancelled: false, answers: { output: ["New sheet"] } };
    })).toEqual({ answers: { output: { answers: ["New sheet"] } } });
  });
  it("maps Claude answers back to question text and keeps its input", async () => {
    const input = { questions: [{ question: "Which columns?", header: "Columns", multiSelect: true,
      options: [{ label: "A" }, { label: "B" }] }] };
    expect(await answerClaudeQuestion(input, async () => ({ cancelled: false, answers: { "0": ["A", "B"] } })))
      .toEqual({ behavior: "allow", updatedInput: { ...input, answers: { "Which columns?": "A, B" } } });
    expect(await answerClaudeQuestion(input, async () => ({ cancelled: true, answers: {} }))).toMatchObject({ behavior: "deny" });
  });
  it("does not invent a native answer when the handler is missing or cancelled", async () => {
    expect(await answerCodexQuestions(request, undefined)).toEqual({ answers: { output: { answers: [] } } });
    const controller = new AbortController();
    controller.abort();
    expect(await answerCodexQuestions(request, async () => { throw new Error("must not ask"); }, controller.signal))
      .toEqual({ answers: { output: { answers: [] } } });
  });
  it("rejects missing, extra, invalid and multiple single-choice answers", () => {
    for (const answers of [{}, { output: ["New sheet"], injected: ["yes"] }, { output: ["Elsewhere"] }, { output: ["New sheet", "Current sheet"] }]) {
      expect(() => validateUserInputResponse(request, { cancelled: false, answers })).toThrow();
    }
  });
  it("returns the pane answer and releases the pending request", async () => {
    const { session, messages } = connected();
    const pending = session.requestUserInput(request);
    const message = messages[0]!;
    expect(message).toMatchObject({ type: "tool_request", tool: "__request_user_input", requestId: "turn" });
    if (message.type !== "tool_request") throw new Error("missing request");
    session.resolveTool(message.callId, { ok: true, data: { cancelled: false, answers: { output: ["New sheet"] } } });
    expect(await pending).toEqual({ cancelled: false, answers: { output: ["New sheet"] } });
    expect(session.pendingToolCount).toBe(0);
    expect(messages.at(-1)).toEqual({ type: "user_input_closed", callId: message.callId });
  });
  it.each(["stop", "disconnect", "end", "provider"])("closes pending questions on %s and ignores late replies", async action => {
    const { session, messages, socket } = connected();
    const controller = new AbortController();
    const pending = session.requestUserInput(request, controller.signal);
    const message = messages[0]!;
    if (message.type !== "tool_request") throw new Error("missing request");
    if (action === "stop") session.abortTurn();
    else if (action === "disconnect") session.detachSocket(socket);
    else if (action === "end") session.endTurn("turn");
    else controller.abort();
    expect(await pending).toEqual({ cancelled: true, answers: {} });
    expect(session.pendingToolCount).toBe(0);
    expect(session.resolveTool(message.callId, { ok: true, data: { answers: { output: ["late"] } } })).toBe(false);
  });
});
