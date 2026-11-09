import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;
afterEach(async () => { await bridge?.close(); bridge = null; });

let httpUrl = "";

async function start(reply: string, adapter: MockAdapter = new MockAdapter({ id: "claude", reply })) {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(adapter) });
  const { url } = await bridge.listen();
  httpUrl = url;
  return `${url.replace("http", "ws")}/ws`;
}

/** A provider whose CLI is missing: the turn fails before it starts. */
class MissingCli extends MockAdapter {
  override async detect() {
    return { ...(await super.detect()), installed: false, hint: "The claude CLI is not installed." };
  }
}

const names = new WeakMap<PaneClient, string>();

/** The bridge asks each pane what its workbook is called; a pane always answers, whatever else it does with messages. */
function answerName(pane: PaneClient, m: { type: string; tool?: string; callId?: string }): void {
  if (m.type === "tool_request" && m.tool === "__workbook_name") {
    pane.send({ type: "tool_result", callId: m.callId!, result: { ok: true, data: { name: names.get(pane) ?? null } } });
  }
}

/** Replace what a pane does with each message (it still gives its name when asked). */
function handle(pane: PaneClient, fn: (m: Parameters<NonNullable<PaneClient["onMessage"]>>[0]) => void): void {
  pane.onMessage = (m) => { answerName(pane, m as never); fn(m); };
}

async function openPane(wsUrl: string, connectionId: string, name: string, options: { approve?: boolean } = {}) {
  const pane = await PaneClient.connect(wsUrl);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: connectionId });
  await pane.waitFor((m) => m.type === "hello_ok");
  pane.send({ type: "workbook_info", name });
  pane.send({ type: "set_provider", provider: "claude" });
  await pane.waitFor((m) => m.type === "provider_status");
  names.set(pane, name);
  pane.onMessage = (m) => {
    answerName(pane, m as never);
    if (m.type === "tool_request" && m.tool === "__authorize_connected_file") {
      pane.send({ type: "tool_result", callId: m.callId, result: { ok: true, data: { approved: options.approve ?? true } } });
    }
  };
  return pane;
}

async function sessionIdOf(pane: PaneClient, name: string): Promise<{ id: string; fileId: string }> {
  await pane.waitFor((m) => m.type === "connected_files" && m.files.some((f) => f.current && f.name === name));
  const session = bridge!.sessions.list().find((item) => item.workbookName === name)!;
  return { id: session.id, fileId: session.fileId };
}

describe("two workbooks talking to each other over the wire", () => {
  it("runs the request as a normal turn in the other file's chat and returns its reply", async () => {
    const wsUrl = await start("Totals written to D2:D40.");
    const budget = await openPane(wsUrl, "budget-connection-1", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-1", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");

    // The Actuals task pane does what main.ts does: run the request as its own turn, tagged with the delegation id.
    handle(actuals, (m) => {
      if (m.type !== "delegated_prompt") return;
      expect(m.from).toEqual({ id: from.fileId, name: "Budget.xlsx", host: "excel" });
      actuals.send({ type: "prompt", requestId: "actuals-turn-1", prompt: m.text, selection: null, model: null, effort: "medium", delegationId: m.delegationId });
    });

    const result = await bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Fill column D with totals" } });
    expect(result).toMatchObject({ ok: true, data: { file: { id: to.fileId, name: "Actuals.xlsx" }, outcome: "completed", reply: "Totals written to D2:D40." } });

    // The turn shows up in the Actuals chat like any other turn.
    await actuals.waitFor((m) => m.type === "turn_finished" && m.requestId === "actuals-turn-1");
    const types = actuals.received.filter((m) => "requestId" in m && m.requestId === "actuals-turn-1").map((m) => m.type);
    expect(types).toContain("message_completed");
    expect(types).toContain("turn_finished");
    // Nothing was written to Budget's own chat.
    expect(budget.received.some((m) => "requestId" in m && m.requestId === "actuals-turn-1")).toBe(false);
    await budget.close(); await actuals.close();
  });

  it("hands the asker the ranges the other file's tools really changed, as links, whatever its reply says", async () => {
    const adapter = new MockAdapter({ id: "claude", reply: "Đã nhập dữ liệu vào Actuals → Sheet9, vùng Z99.", toolCalls: [
      { tool: "write_range", args: { sheet: { name: "Data" }, address: "A1", values: [["Sản phẩm", "SL"], ["Táo", 3]] } },
      { tool: "format_range", args: { sheet: { name: "Data" }, address: "A1:B1", format: { bold: true } } },
    ] });
    const wsUrl = await start("unused", adapter);
    const budget = await openPane(wsUrl, "budget-connection-1", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-1", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    handle(actuals, (m) => {
      if (m.type === "delegated_prompt") actuals.send({ type: "prompt", requestId: "actuals-turn-1", prompt: m.text, selection: null, model: null, effort: "medium", delegationId: m.delegationId });
      // What the Excel executor answers: the sheet and the range it really wrote.
      if (m.type === "tool_request" && (m.tool === "write_range" || m.tool === "format_range")) {
        const address = m.tool === "write_range" ? "A1:B2" : "A1:B1";
        actuals.send({ type: "tool_result", callId: m.callId, result: { ok: true, data: { sheet: "Data", address, rowsWritten: 2, columnsWritten: 2 } } });
      }
    });

    const result = await bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Fill a sample sales table" } });
    expect(result).toMatchObject({ ok: true, data: { outcome: "completed", reply: "Đã nhập dữ liệu vào Actuals → Sheet9, vùng Z99." } });
    // The reply names the wrong place (Sheet9, Z99): the links come from what the tools did.
    expect((result as { data: { changed: unknown[] } }).data.changed).toEqual([
      { file: "Actuals.xlsx", sheet: "Data", address: "A1:B2", link: "[A1:B2](cell:?file=Actuals.xlsx&sheet=Data&range=A1:B2)" },
      { file: "Actuals.xlsx", sheet: "Data", address: "A1:B1", link: "[A1:B1](cell:?file=Actuals.xlsx&sheet=Data&range=A1:B1)" },
    ]);
    await budget.close(); await actuals.close();
  });

  it("tells the asker when the other pane declines because it is busy", async () => {
    const wsUrl = await start("unused");
    const budget = await openPane(wsUrl, "budget-connection-2", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-2", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    handle(actuals, (m) => {
      if (m.type === "delegated_prompt") actuals.send({ type: "delegation_declined", delegationId: m.delegationId, reason: "That file is busy with its own request." });
    });
    const result = await bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Do something" } });
    expect(result).toMatchObject({ ok: false, error: { code: "unsupported", message: "[Actuals.xlsx] That file is busy with its own request." } });
    await budget.close(); await actuals.close();
  });

  it("sends nothing to the other pane when the user declines", async () => {
    const wsUrl = await start("unused");
    const budget = await openPane(wsUrl, "budget-connection-3", "Budget.xlsx", { approve: false });
    const actuals = await openPane(wsUrl, "actuals-connection-3", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    const result = await bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Do something" } });
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(actuals.received.some((m) => m.type === "delegated_prompt")).toBe(false);
    await budget.close(); await actuals.close();
  });

  it("ends the request when the other pane closes while working", async () => {
    const wsUrl = await start("unused");
    const budget = await openPane(wsUrl, "budget-connection-4", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-4", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    handle(actuals, (m) => { if (m.type === "delegated_prompt") void actuals.close(); });
    const result = await bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Do something" } });
    expect(result).toMatchObject({ ok: false, error: { code: "disconnected" } });
    await budget.close();
  });
});

describe("a request nobody waits for, and a turn that fails", () => {
  const prompt = (pane: PaneClient, delegationId: string, requestId: string) =>
    pane.send({ type: "prompt", requestId, prompt: "Fill column D", selection: null, model: null, effort: "medium", delegationId });
  const ran = (pane: PaneClient, requestId: string) => pane.received.some((m) => "requestId" in m && m.requestId === requestId && m.type === "message_completed");

  it("does not run a prompt whose delegation is unknown, and ends it as cancelled", async () => {
    const wsUrl = await start("should not run");
    const actuals = await openPane(wsUrl, "actuals-connection-5", "Actuals.xlsx");
    prompt(actuals, "never-asked", "stale-1");
    await actuals.waitFor((m) => m.type === "turn_finished" && m.requestId === "stale-1");
    expect(actuals.received.find((m) => m.type === "error" && m.requestId === "stale-1")).toMatchObject({ code: "cancelled" });
    expect(actuals.received.find((m) => m.type === "turn_finished" && m.requestId === "stale-1")).toMatchObject({ reason: "cancelled" });
    expect(ran(actuals, "stale-1")).toBe(false);
    await actuals.close();
  });

  it("does not run the request when the asker stopped before the other pane answered", async () => {
    const wsUrl = await start("should not run");
    const budget = await openPane(wsUrl, "budget-connection-6", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-6", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    const controller = new AbortController();
    const result = bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Fill column D" }, signal: controller.signal });
    const delivered = await actuals.waitFor((m) => m.type === "delegated_prompt") as { delegationId: string };
    controller.abort();
    expect(await result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    // The pane only now answers, as a slow pane would.
    prompt(actuals, delivered.delegationId, "late-1");
    await actuals.waitFor((m) => m.type === "turn_finished" && m.requestId === "late-1");
    expect(ran(actuals, "late-1")).toBe(false);
    await budget.close(); await actuals.close();
  });

  it("stops waiting for the other file when the caller hangs up, instead of letting the request run on", async () => {
    const wsUrl = await start("should not run");
    const budget = await openPane(wsUrl, "budget-connection-7", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-7", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    const token = bridge!.sessions.list().find((item) => item.id === from.id)!.sessionToken;
    const hangUp = new AbortController();
    const call = fetch(`${httpUrl}/internal/tool`, {
      method: "POST",
      headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
      body: JSON.stringify({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Fill column D" } }),
      signal: hangUp.signal,
    }).catch(() => null);
    const delivered = await actuals.waitFor((m) => m.type === "delegated_prompt") as { delegationId: string };
    hangUp.abort();
    await call;
    // Give the bridge a moment to notice the closed connection, then let the pane answer.
    await new Promise((resolve) => setTimeout(resolve, 100));
    prompt(actuals, delivered.delegationId, "orphan-1");
    await actuals.waitFor((m) => m.type === "turn_finished" && m.requestId === "orphan-1");
    expect(ran(actuals, "orphan-1")).toBe(false);
    await budget.close(); await actuals.close();
  });

  it("tells the asker the other file failed, with the reason, when its turn errors before it starts", async () => {
    const wsUrl = await start("unused", new MissingCli({ id: "claude" }));
    const budget = await openPane(wsUrl, "budget-connection-8", "Budget.xlsx");
    const actuals = await openPane(wsUrl, "actuals-connection-8", "Actuals.xlsx");
    const from = await sessionIdOf(budget, "Budget.xlsx");
    const to = await sessionIdOf(actuals, "Actuals.xlsx");
    handle(actuals, (m) => { if (m.type === "delegated_prompt") prompt(actuals, m.delegationId, "failing-1"); });
    const result = await bridge!.gateway.invoke({ sessionId: from.id, tool: "send_to_connected_file", args: { file: to.fileId, message: "Fill column D" } });
    expect(result).toMatchObject({ ok: false, error: { code: "internal_error", message: expect.stringContaining("The claude CLI is not installed.") } });
    await budget.close(); await actuals.close();
  });
});
