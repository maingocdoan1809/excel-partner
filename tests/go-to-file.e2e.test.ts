import { afterEach, describe, expect, it, vi } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { ToolResult } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;
const panes: PaneClient[] = [];
afterEach(async () => {
  for (const pane of panes.splice(0)) await pane.close();
  await bridge?.close();
  bridge = null;
});

const activate = vi.fn(async (_name: string) => undefined);

async function start(options: { activateWorkbook?: (name: string) => Promise<unknown> } = {}) {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, activateWorkbook: options.activateWorkbook ?? activate,
    registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", reply: "ok" })) });
  return (await bridge.listen()).url;
}

interface OpenPane { pane: PaneClient; token: string; shown: Array<{ sheet?: string; address: string }> }

/** A task pane for a workbook called `name`; `answer` is how it responds to a request to show a range. */
async function openPane(url: string, name: string, answer: (args: { sheet?: string; address: string }) => ToolResult = () => ({ ok: true, data: {} })): Promise<OpenPane> {
  const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  panes.push(pane);
  const shown: OpenPane["shown"] = [];
  pane.onMessage = (message) => {
    if (message.type !== "tool_request") return;
    if (message.tool === "__workbook_name") pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { name } } });
    if (message.tool === "__go_to") {
      const args = message.args as { sheet?: string; address: string };
      shown.push(args);
      pane.send({ type: "tool_result", callId: message.callId, result: answer(args) });
    }
  };
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: `connection-${name}-${panes.length}` });
  const hello = await pane.waitFor((m) => m.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("no handshake");
  pane.send({ type: "workbook_info", name });
  await pane.waitFor((m) => m.type === "connected_files" && m.files.some((f) => f.name === name && f.current));
  return { pane, token: hello.sessionToken, shown };
}

const goTo = (url: string, token: string | null, body: unknown) => fetch(`${url}/session/go-to`, {
  method: "POST",
  headers: { "content-type": "application/json", ...(token ? { authorization: `Bearer ${token}` } : {}) },
  body: JSON.stringify(body),
});

describe("following a link into another open file", () => {
  it("asks that file's pane to show the range, finding the file by the name the text used, and brings its window forward", async () => {
    activate.mockClear();
    const url = await start();
    const budget = await openPane(url, "Budget.xlsx");
    const book2 = await openPane(url, "Book2.xlsx");
    const answer = await goTo(url, budget.token, { file: "Book2", sheet: "Sheet1", address: "A1:E8" });
    expect(answer.status).toBe(200);
    expect(await answer.json()).toEqual({ ok: true });
    expect(book2.shown).toEqual([{ sheet: "Sheet1", address: "A1:E8" }]);
    expect(budget.shown).toEqual([]);
    expect(activate).toHaveBeenCalledWith("Book2.xlsx");
  });

  it("leaves out the sheet when the link named none, and still finds the file with its extension", async () => {
    const url = await start();
    const budget = await openPane(url, "Budget.xlsx");
    const book2 = await openPane(url, "Book2.xlsx");
    expect(await (await goTo(url, budget.token, { file: "book2.XLSX", address: "$B$2" })).json()).toEqual({ ok: true });
    expect(book2.shown).toEqual([{ address: "$B$2" }]);
  });

  it("says so when the file is not open, and never treats the asking file as the other one", async () => {
    const url = await start();
    const budget = await openPane(url, "Budget.xlsx");
    await openPane(url, "Book2.xlsx");
    expect(await (await goTo(url, budget.token, { file: "Missing.xlsx", address: "A1" })).json())
      .toEqual({ ok: false, error: '"Missing.xlsx" is not open in Excel with this add-in.' });
    expect(await (await goTo(url, budget.token, { file: "Budget", address: "A1" })).json()).toMatchObject({ ok: false });
    expect(budget.shown).toEqual([]);
  });

  it("does not pick one of two files with the same name", async () => {
    const url = await start();
    const asker = await openPane(url, "Budget.xlsx");
    const first = await openPane(url, "Same.xlsx");
    const second = await openPane(url, "Same.xlsx");
    const answer = await (await goTo(url, asker.token, { file: "Same", address: "A1" })).json() as { ok: boolean; error: string };
    expect(answer.ok).toBe(false);
    expect(answer.error).toContain("More than one open file");
    expect([first.shown, second.shown]).toEqual([[], []]);
  });

  it("passes on why the other file could not show it, and does not bring its window forward", async () => {
    activate.mockClear();
    const url = await start();
    const budget = await openPane(url, "Budget.xlsx");
    await openPane(url, "Book2.xlsx", () => ({ ok: false, error: { code: "office_error", message: 'There is no sheet named "Nope" in this workbook.' } }));
    expect(await (await goTo(url, budget.token, { file: "Book2", sheet: "Nope", address: "A1" })).json())
      .toEqual({ ok: false, error: '[Book2.xlsx] There is no sheet named "Nope" in this workbook.' });
    expect(activate).not.toHaveBeenCalled();
  });

  it("still succeeds when the window cannot be brought forward", async () => {
    const url = await start({ activateWorkbook: async () => { throw new Error("no Excel window"); } });
    const budget = await openPane(url, "Budget.xlsx");
    const book2 = await openPane(url, "Book2.xlsx");
    expect(await (await goTo(url, budget.token, { file: "Book2", address: "A1" })).json()).toEqual({ ok: true });
    expect(book2.shown).toHaveLength(1);
  });

  it("needs a connected pane's token and a sane target", async () => {
    const url = await start();
    const budget = await openPane(url, "Budget.xlsx");
    await openPane(url, "Book2.xlsx");
    expect((await goTo(url, null, { file: "Book2", address: "A1" })).status).toBe(401);
    expect((await goTo(url, "not-a-token", { file: "Book2", address: "A1" })).status).toBe(401);
    for (const body of [{}, { file: "Book2" }, { address: "A1" }, { file: "Book2", address: "A1;rm" }, { file: "Book2", address: "A1", sheet: 5 }, { file: "x".repeat(300), address: "A1" }]) {
      expect((await goTo(url, budget.token, body)).status, JSON.stringify(body)).toBe(400);
    }
    const raw = await fetch(`${url}/session/go-to`, { method: "POST", headers: { authorization: `Bearer ${budget.token}` }, body: "{not json" });
    expect(raw.status).toBe(400);
  });

  it("works from a browser tab, which is no file itself", async () => {
    const url = await start();
    const book2 = await openPane(url, "Book2.xlsx");
    const tab = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    panes.push(tab);
    tab.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "browser-tab-0001", clientInfo: { platform: "Browser", standalone: true } });
    const hello = await tab.waitFor((m) => m.type === "hello_ok");
    if (hello.type !== "hello_ok") throw new Error("no handshake");
    expect(await (await goTo(url, hello.sessionToken, { file: "Book2", sheet: "Sheet1", address: "A1" })).json()).toEqual({ ok: true });
    expect(book2.shown).toEqual([{ sheet: "Sheet1", address: "A1" }]);
  });
});
