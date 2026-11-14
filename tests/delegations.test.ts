import { describe, expect, it, vi } from "vitest";
import type { WebSocket } from "ws";
import { Delegations } from "../bridge/src/delegations.js";
import { WorkbookSessions } from "../bridge/src/workbook-sessions.js";
import type { WorkbookSession } from "../bridge/src/session.js";
import type { ProviderEvent, ServerMessage } from "../protocol/src/index.js";

function connected(sessions: WorkbookSessions, id: string) {
  const session = sessions.create(id, null);
  const sent: ServerMessage[] = [];
  session.attachSocket({ readyState: 1, close() {}, send: (raw: string) => sent.push(JSON.parse(raw)) } as unknown as WebSocket);
  return { session, sent };
}

function setup(limits = { startMs: 1000, runMs: 1000 }) {
  const sessions = new WorkbookSessions({ ttlMs: 60_000 });
  const cancelled: Array<[string, string]> = [];
  const delegations = new Delegations((session: WorkbookSession, requestId) => cancelled.push([session.fileId, requestId]), limits);
  const target = connected(sessions, "target");
  const delivered = () => target.sent.find((m) => m.type === "delegated_prompt") as Extract<ServerMessage, { type: "delegated_prompt" }>;
  return { sessions, delegations, target, cancelled, delivered };
}

const from = { id: "f9", name: "Budget.xlsx" };
const text = (message: string): ProviderEvent => ({ type: "message_completed", requestId: "r1", text: message });

describe("Delegations", () => {
  it("delivers the request, follows the turn and returns the last message", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "Total column D");
    const { delegationId } = delivered();
    expect(delivered()).toMatchObject({ from, text: "Total column D" });
    expect(delegations.accept(target.session, delegationId, "r1")).toBe(true);
    delegations.observe(target.session, "r1", text("Working on it."));
    delegations.observe(target.session, "r1", text("D2:D40 now hold the totals."));
    delegations.finish(target.session, "r1", "completed");
    expect(await result).toMatchObject({ ok: true, data: { outcome: "completed", reply: "D2:D40 now hold the totals." } });
    expect(delegations.size).toBe(0);
  });

  it("reports an error turn as a failure with the error message, not as finished work", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "x");
    delegations.accept(target.session, delivered().delegationId, "r1");
    delegations.observe(target.session, "r1", { type: "error", requestId: "r1", code: "provider_error", message: "model overloaded" });
    delegations.finish(target.session, "r1", "error");
    expect(await result).toMatchObject({ ok: false, error: { code: "internal_error", message: expect.stringContaining("model overloaded") } });
  });

  it("still reports a failure, and what was said so far, when the turn gave no error text", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "x");
    delegations.accept(target.session, delivered().delegationId, "r1");
    delegations.observe(target.session, "r1", text("Half done."));
    delegations.finish(target.session, "r1", "error");
    expect(await result).toMatchObject({ ok: false, error: { code: "internal_error", message: expect.stringContaining("no details"), detail: { reply: "Half done." } } });
  });

  it("reports a turn stopped in the other window as cancelled, not as a finished request", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "x");
    delegations.accept(target.session, delivered().delegationId, "r1");
    delegations.finish(target.session, "r1", "cancelled");
    expect(await result).toMatchObject({ ok: false, error: { code: "cancelled", message: expect.stringContaining("stopped") } });
  });

  it("only accepts the pane the request was sent to, and only once", () => {
    const { sessions, delegations, target, delivered } = setup();
    const other = connected(sessions, "other");
    void delegations.run(target.session, from, "x");
    const { delegationId } = delivered();
    expect(delegations.accept(other.session, delegationId, "r1")).toBe(false);
    expect(delegations.accept(target.session, delegationId, "r1")).toBe(true);
    expect(delegations.accept(target.session, delegationId, "r2")).toBe(false);
    expect(delegations.accept(target.session, "unknown", "r3")).toBe(false);
  });

  it("fails when the pane declines", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "x");
    delegations.decline(target.session, delivered().delegationId, "That file is busy with its own request.");
    expect(await result).toMatchObject({ ok: false, error: { code: "unsupported", message: "That file is busy with its own request." } });
  });

  it("fails when the pane never starts the request", async () => {
    const { delegations, target } = setup({ startMs: 20, runMs: 1000 });
    expect(await delegations.run(target.session, from, "x")).toMatchObject({ ok: false, error: { code: "timeout" } });
    expect(delegations.size).toBe(0);
  });

  it("stops a request that runs too long", async () => {
    const { delegations, target, cancelled, delivered } = setup({ startMs: 1000, runMs: 20 });
    const result = delegations.run(target.session, from, "x");
    delegations.accept(target.session, delivered().delegationId, "r1");
    expect(await result).toMatchObject({ ok: false, error: { code: "timeout" } });
    expect(cancelled).toEqual([["f1", "r1"]]);
  });

  it("stops the other file's turn when the asker cancels", async () => {
    const { delegations, target, cancelled, delivered } = setup();
    const controller = new AbortController();
    const result = delegations.run(target.session, from, "x", controller.signal);
    delegations.accept(target.session, delivered().delegationId, "r1");
    controller.abort();
    expect(await result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(cancelled).toEqual([["f1", "r1"]]);
  });

  it("does nothing when already cancelled before it starts", async () => {
    const { delegations, target } = setup();
    const controller = new AbortController();
    controller.abort();
    expect(await delegations.run(target.session, from, "x", controller.signal)).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(target.sent).toEqual([]);
  });

  it("ends the request when the other pane closes, and when it is not connected", async () => {
    const { sessions, delegations, target } = setup();
    const waiting = delegations.run(target.session, from, "x");
    delegations.sessionClosed(target.session);
    expect(await waiting).toMatchObject({ ok: false, error: { code: "disconnected" } });
    const gone = connected(sessions, "gone");
    gone.session.detachSocket(gone.session.socket!);
    expect(await delegations.run(gone.session, from, "x")).toMatchObject({ ok: false, error: { code: "disconnected" } });
  });

  it("ignores events of turns it does not follow and never settles twice", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "x");
    delegations.accept(target.session, delivered().delegationId, "r1");
    delegations.observe(target.session, "someone-else", text("not mine"));
    delegations.finish(target.session, "someone-else", "completed");
    expect(delegations.size).toBe(1);
    delegations.finish(target.session, "r1", "completed");
    delegations.finish(target.session, "r1", "error");
    expect(await result).toMatchObject({ ok: true, data: { outcome: "completed" } });
  });

  it("cuts a very long reply", async () => {
    const { delegations, target, delivered } = setup();
    const result = delegations.run(target.session, from, "x");
    delegations.accept(target.session, delivered().delegationId, "r1");
    delegations.observe(target.session, "r1", text("a".repeat(50_000)));
    delegations.finish(target.session, "r1", "completed");
    expect(((await result) as { data: { reply: string } }).data.reply).toHaveLength(12_000);
  });
});

describe("Delegations: what the other file really changed", () => {
  const wrote = (sheet: string, address: string) => ({ ok: true as const, data: { sheet, address, rowsWritten: 1, columnsWritten: 1 } });

  /** A request that is running in the target's turn r1, so that its tools count for it. */
  function running(name: string | null = "Book2.xlsx") {
    const world = setup();
    world.target.session.workbookName = name;
    const result = world.delegations.run(world.target.session, from, "fill A1:E8");
    world.delegations.accept(world.target.session, world.delivered().delegationId, "r1");
    world.target.session.beginTurn("r1");
    return { ...world, result };
  }

  it("hands back the ranges its tools changed, with links ready to paste, read from the tool results", async () => {
    const { delegations, target, result } = running();
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "A1:E8"));
    delegations.recordChange(target.session, "format_range", wrote("Sheet1", "A1:E1"));
    delegations.recordChange(target.session, "set_formula", wrote("Tổng hợp", "B10"));
    delegations.observe(target.session, "r1", text("Đã nhập dữ liệu vào Book2 → Sheet1, vùng A1:E8."));
    delegations.finish(target.session, "r1", "completed");
    const data = ((await result) as { data: { changed: unknown[]; changedNote: string } }).data;
    expect(data.changed).toEqual([
      { file: "Book2.xlsx", sheet: "Sheet1", address: "A1:E8", link: "[A1:E8](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E8)" },
      { file: "Book2.xlsx", sheet: "Sheet1", address: "A1:E1", link: "[A1:E1](cell:?file=Book2.xlsx&sheet=Sheet1&range=A1:E1)" },
      { file: "Book2.xlsx", sheet: "Tổng hợp", address: "B10", link: "[B10](cell:?file=Book2.xlsx&sheet=T%E1%BB%95ng%20h%E1%BB%A3p&range=B10)" },
    ]);
    expect(data.changedNote).toContain("paste these links");
  });

  it("counts a range once, and only what a successful tool that names its range changed", async () => {
    const { delegations, target, result } = running();
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "A1"));
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "A1"));
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "$A$1"));
    delegations.recordChange(target.session, "read_range", wrote("Sheet1", "B2"));
    delegations.recordChange(target.session, "write_range", { ok: false, error: { code: "office_error", message: "no" } });
    delegations.recordChange(target.session, "write_range", { ok: true, data: { rowsWritten: 1 } });
    delegations.recordChange(target.session, "write_range", wrote("", "B2"));
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "not a range"));
    delegations.finish(target.session, "r1", "completed");
    expect(((await result) as { data: { changed: Array<{ address: string }> } }).data.changed.map((c) => c.address)).toEqual(["A1"]);
  });

  it("keeps at most thirty ranges", async () => {
    const { delegations, target, result } = running();
    for (let row = 1; row <= 40; row++) delegations.recordChange(target.session, "write_range", wrote("S", `A${row}`));
    delegations.finish(target.session, "r1", "completed");
    expect(((await result) as { data: { changed: unknown[] } }).data.changed).toHaveLength(30);
  });

  it("ignores tools of another turn, of another pane and of no request, and says nothing when nothing changed", async () => {
    const { delegations, target, sessions, result } = running();
    const other = connected(sessions, "other");
    other.session.beginTurn("r1");
    delegations.recordChange(other.session, "write_range", wrote("Sheet1", "A1"));
    target.session.endTurn("r1");
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "A2"));
    target.session.beginTurn("r2");
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "A3"));
    delegations.finish(target.session, "r1", "completed");
    const data = ((await result) as { data: Record<string, unknown> }).data;
    expect(data).not.toHaveProperty("changed");
    expect(data).not.toHaveProperty("changedNote");
    expect(data.reply).toBeDefined();
  });

  it("gives no links for a file that has no name yet: nothing could find it from another chat", async () => {
    const { delegations, target, result } = running(null);
    delegations.recordChange(target.session, "write_range", wrote("Sheet1", "A1:E8"));
    delegations.finish(target.session, "r1", "completed");
    expect(((await result) as { data: Record<string, unknown> }).data).not.toHaveProperty("changed");
  });
});

vi.useRealTimers();
