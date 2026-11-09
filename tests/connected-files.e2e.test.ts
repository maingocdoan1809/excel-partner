import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { ConnectedFile } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;
afterEach(async () => { await bridge?.close(); bridge = null; });

async function start() {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", reply: "ok" })) });
  const { url } = await bridge.listen();
  return url;
}

async function openPane(url: string, connectionId: string, name?: string) {
  const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: connectionId });
  await pane.waitFor((m) => m.type === "hello_ok");
  if (name) pane.send({ type: "workbook_info", name });
  return pane;
}

const lastFiles = (pane: PaneClient): ConnectedFile[] => {
  const messages = pane.received.filter((m) => m.type === "connected_files");
  const last = messages[messages.length - 1];
  return last && last.type === "connected_files" ? last.files : [];
};

describe("connected files over the wire", () => {
  it("tells each task pane which workbooks are open, itself first, and follows names and closing", async () => {
    const url = await start();
    const budget = await openPane(url, "budget-connection-1", "Budget.xlsx");
    await budget.waitFor((m) => m.type === "connected_files" && m.files[0]?.name === "Budget.xlsx");
    const actuals = await openPane(url, "actuals-connection-1", "Actuals.xlsx");

    await actuals.waitFor((m) => m.type === "connected_files" && m.files.length === 2 && m.files.every((f) => f.name));
    expect(lastFiles(actuals)).toEqual([
      { id: "f2", name: "Actuals.xlsx", host: "excel", current: true },
      { id: "f1", name: "Budget.xlsx", host: "excel", current: false },
    ]);
    await budget.waitFor((m) => m.type === "connected_files" && m.files.length === 2 && m.files.every((f) => f.name));
    expect(lastFiles(budget)[0]).toEqual({ id: "f1", name: "Budget.xlsx", host: "excel", current: true });

    budget.send({ type: "workbook_info", name: "Budget (final).xlsx" });
    await actuals.waitFor((m) => m.type === "connected_files" && m.files.some((f) => f.name === "Budget (final).xlsx"));

    await budget.close();
    await actuals.waitFor((m) => m.type === "connected_files" && m.files.length === 1);
    expect(lastFiles(actuals)).toEqual([{ id: "f2", name: "Actuals.xlsx", host: "excel", current: true }]);
    await actuals.close();
  });

  it("ignores a workbook_info without a usable name", async () => {
    const url = await start();
    const pane = await openPane(url, "only-connection-01", "Real.xlsx");
    await pane.waitFor((m) => m.type === "connected_files" && m.files[0]?.name === "Real.xlsx");
    pane.send({ type: "workbook_info", name: "   " });
    pane.send({ type: "workbook_info", name: 42 as unknown as string });
    pane.send({ type: "workbook_info", name: "Next.xlsx" });
    await pane.waitFor((m) => m.type === "connected_files" && m.files[0]?.name === "Next.xlsx");
    expect(pane.received.filter((m) => m.type === "connected_files" && m.files[0]?.name === "")).toHaveLength(1);
    await pane.close();
  });
});

describe("names are read when someone looks, not on a timer", () => {
  /** A pane whose workbook is called `current.name`; it counts how often the bridge asks. */
  function namedPane(pane: PaneClient, current: { name: string | null }) {
    const asked = { count: 0 };
    pane.onMessage = (m) => {
      if (m.type === "tool_request" && m.tool === "__workbook_name") {
        asked.count++;
        pane.send({ type: "tool_result", callId: m.callId, result: { ok: true, data: { name: current.name } } });
      }
    };
    return asked;
  }

  it("shows every pane the new name after a rename, once one pane asks for the list", async () => {
    const url = await start();
    const budgetName = { name: "Budget.xlsx" };
    const budget = await openPane(url, "budget-connection-1", "Budget.xlsx");
    namedPane(budget, budgetName);
    const actuals = await openPane(url, "actuals-connection-1", "Actuals.xlsx");
    namedPane(actuals, { name: "Actuals.xlsx" });
    await actuals.waitFor((m) => m.type === "connected_files" && m.files.length === 2 && m.files.every((f) => f.name));

    // Save As in the Budget window: Excel gives no event, and the bridge still has the old name.
    budgetName.name = "Budget (final).xlsx";
    expect(lastFiles(actuals).map((f) => f.name)).toContain("Budget.xlsx");

    actuals.send({ type: "refresh_connected_files" });
    await actuals.waitFor((m) => m.type === "connected_files" && m.files.some((f) => f.name === "Budget (final).xlsx"));
    await budget.waitFor((m) => m.type === "connected_files" && m.files.some((f) => f.name === "Budget (final).xlsx"));
    expect(lastFiles(actuals).map((f) => f.name)).toEqual(["Actuals.xlsx", "Budget (final).xlsx"]);
    await budget.close(); await actuals.close();
  });

  it("asks nobody while nobody looks", async () => {
    const url = await start();
    const asked = namedPane(await openPane(url, "idle-connection-01", "Idle.xlsx"), { name: "Idle.xlsx" });
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(asked.count).toBe(0);
  });

  it("asks each pane once for requests that arrive together", async () => {
    const url = await start();
    const first = await openPane(url, "burst-connection-01", "A.xlsx");
    const askedFirst = namedPane(first, { name: "A.xlsx" });
    const second = await openPane(url, "burst-connection-02", "B.xlsx");
    const askedSecond = namedPane(second, { name: "B.xlsx" });
    await second.waitFor((m) => m.type === "connected_files" && m.files.length === 2);
    const before = second.received.filter((m) => m.type === "connected_files").length;

    first.send({ type: "refresh_connected_files" });
    second.send({ type: "refresh_connected_files" });
    second.send({ type: "refresh_connected_files" });
    await second.waitFor((m, i) => m.type === "connected_files" && second.received.filter((x) => x.type === "connected_files").length > before);
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(askedFirst.count).toBe(1);
    expect(askedSecond.count).toBe(1);
    await first.close(); await second.close();
  });

  it("keeps the name it has when a pane cannot say", async () => {
    const url = await start();
    const pane = await openPane(url, "silent-connection-1", "Kept.xlsx");
    namedPane(pane, { name: null });
    pane.send({ type: "refresh_connected_files" });
    await pane.waitFor((m, i) => m.type === "connected_files" && i > 1);
    expect(lastFiles(pane)).toEqual([{ id: "f1", name: "Kept.xlsx", host: "excel", current: true }]);
    await pane.close();
  });
});
