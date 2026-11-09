import { describe, expect, it } from "vitest";
import type { WebSocket } from "ws";
import { Delegations } from "../bridge/src/delegations.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import type { WorkbookSession } from "../bridge/src/session.js";
import { WORKBOOK_NAME_REQUEST } from "../protocol/src/index.js";

const limits = { maxReadCells: 100, maxWriteCells: 100, maxSheets: 10, maxSelectionCells: 10 };

/** Source pane: answers approvals. Target pane: runs a delegated request and ends it with `reply`. */
function setup(options: { approve?: boolean; reply?: string; targetProvider?: "claude" | null; sourceHost?: "excel" | "word" | "powerpoint"; targetHost?: "excel" | "word" | "powerpoint" } = {}) {
  const sessions = new SessionManager({ ttlMs: 60_000 });
  const source = sessions.create("source", options.sourceHost ? { host: options.sourceHost } : null);
  const target = sessions.create("target", options.targetHost ? { host: options.targetHost } : null);
  source.workbookName = "Budget.xlsx";
  target.workbookName = "Actuals.xlsx";
  target.provider = options.targetProvider === undefined ? "claude" : options.targetProvider;
  const approvals: unknown[] = [];
  const delivered: unknown[] = [];
  const delegations = new Delegations(() => {}, { startMs: 500, runMs: 500 });
  source.attachSocket({
    readyState: 1, close() {},
    send(raw: string) {
      const message = JSON.parse(raw);
      if (message.type !== "tool_request") return;
      if (message.tool === WORKBOOK_NAME_REQUEST) return void queueMicrotask(() => source.resolveTool(message.callId, { ok: true, data: { name: source.workbookName } }));
      approvals.push({ tool: message.tool, args: message.args });
      queueMicrotask(() => source.resolveTool(message.callId, { ok: true, data: { approved: options.approve ?? true } }));
    },
  } as unknown as WebSocket);
  target.attachSocket({
    readyState: 1, close() {},
    send(raw: string) {
      const message = JSON.parse(raw);
      if (message.type === "tool_request" && message.tool === WORKBOOK_NAME_REQUEST) return void queueMicrotask(() => target.resolveTool(message.callId, { ok: true, data: { name: target.workbookName } }));
      if (message.type !== "delegated_prompt") return;
      delivered.push(message);
      queueMicrotask(() => {
        delegations.accept(target, message.delegationId, "target-turn");
        delegations.observe(target, "target-turn", { type: "message_completed", requestId: "target-turn", text: options.reply ?? "Done in Actuals." });
        delegations.finish(target, "target-turn", "completed");
      });
    },
  } as unknown as WebSocket);
  const gateway = new ToolGateway(sessions, limits, 1000, undefined, false, null, null, null, undefined, null, null, null, delegations);
  return { sessions, source, target, gateway, approvals, delivered };
}

const send = (gateway: ToolGateway, session: WorkbookSession, file: string, message = "Fill column D with totals") =>
  gateway.invoke({ sessionId: session.id, tool: "send_to_connected_file", args: { file, message } });

describe("send_to_connected_file", () => {
  it("sends a request between files of different Office apps, naming the kind of file to the person", async () => {
    const { gateway, source, target, approvals, delivered } = setup({ sourceHost: "word", targetHost: "powerpoint", reply: "Slide 3 now shows the figure." });
    source.workbookName = "Plan.docx";
    target.workbookName = "Review.pptx";
    source.beginTurn("req-1");
    const result = await send(gateway, source, target.fileId, "Put the revenue figure on slide 3");
    expect(approvals).toEqual([{ tool: "__authorize_connected_file", args: { fileId: "f2", name: "Review.pptx", host: "powerpoint", action: "message", text: "Put the revenue figure on slide 3" } }]);
    expect(delivered).toEqual([expect.objectContaining({ from: { id: "f1", name: "Plan.docx", host: "word" }, text: "Put the revenue figure on slide 3" })]);
    expect(result).toMatchObject({ ok: true, data: { file: { id: "f2", name: "Review.pptx", host: "powerpoint" }, outcome: "completed", reply: "Slide 3 now shows the figure." } });
  });

  it("asks the user with the exact text, delivers it with the sender's name and returns the other assistant's reply", async () => {
    const { gateway, source, target, approvals, delivered } = setup({ reply: "D2:D40 now hold the totals." });
    source.beginTurn("req-1");
    const result = await send(gateway, source, target.fileId);
    expect(approvals).toEqual([{ tool: "__authorize_connected_file", args: { fileId: "f2", name: "Actuals.xlsx", host: "excel", action: "message", text: "Fill column D with totals" } }]);
    expect(delivered).toEqual([expect.objectContaining({ type: "delegated_prompt", from: { id: "f1", name: "Budget.xlsx", host: "excel" }, text: "Fill column D with totals" })]);
    expect(result).toMatchObject({ ok: true, data: { file: { id: "f2", name: "Actuals.xlsx" }, outcome: "completed", reply: "D2:D40 now hold the totals." } });
  });

  it("asks again for every message, because what the other assistant does is not limited to reading", async () => {
    const { gateway, source, target, approvals } = setup();
    source.beginTurn("req-1");
    await send(gateway, source, target.fileId, "first");
    await send(gateway, source, target.fileId, "second");
    expect(approvals).toHaveLength(2);
  });

  it("sends nothing when the user declines", async () => {
    const { gateway, source, target, delivered } = setup({ approve: false });
    source.beginTurn("req-1");
    expect(await send(gateway, source, target.fileId)).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(delivered).toEqual([]);
  });

  it("refuses the current file, an unknown id and a file without a provider, before asking anyone", async () => {
    const { gateway, source, target, approvals, delivered } = setup({ targetProvider: null });
    source.beginTurn("req-1");
    expect(await send(gateway, source, source.fileId)).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await send(gateway, source, "f99")).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(await send(gateway, source, target.fileId)).toMatchObject({ ok: false, error: { code: "unsupported", message: expect.stringContaining("no AI provider") } });
    expect(approvals).toEqual([]);
    expect(delivered).toEqual([]);
  });

  it("refuses a file that is busy with its own request", async () => {
    const { gateway, source, target, approvals, delivered } = setup();
    source.beginTurn("req-1");
    target.beginTurn("its-own");
    expect(await send(gateway, source, target.fileId)).toMatchObject({ ok: false, error: { code: "unsupported", message: expect.stringContaining("busy") } });
    expect(approvals).toEqual([]);
    expect(delivered).toEqual([]);
  });

  it("names the file when the other file fails", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const source = sessions.create("a", null);
    const target = sessions.create("b", null);
    target.workbookName = "Actuals.xlsx";
    target.provider = "claude";
    source.attachSocket({ readyState: 1, close() {}, send(raw: string) { const m = JSON.parse(raw); queueMicrotask(() => source.resolveTool(m.callId, { ok: true, data: m.tool === WORKBOOK_NAME_REQUEST ? { name: source.workbookName } : { approved: true } })); } } as unknown as WebSocket);
    // A pane that gives its name but never starts the request.
    target.attachSocket({ readyState: 1, close() {}, send(raw: string) { const m = JSON.parse(raw); if (m.tool === WORKBOOK_NAME_REQUEST) queueMicrotask(() => target.resolveTool(m.callId, { ok: true, data: { name: target.workbookName } })); } } as unknown as WebSocket);
    const gateway = new ToolGateway(sessions, limits, 1000, undefined, false, null, null, null, undefined, null, null, null, new Delegations(() => {}, { startMs: 30, runMs: 30 }));
    source.beginTurn("req-1");
    expect(await send(gateway, source, target.fileId)).toMatchObject({ ok: false, error: { code: "timeout", message: expect.stringMatching(/^\[Actuals\.xlsx\] /) } });
  });
});
