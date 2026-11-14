// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { BridgeClient } from "../frontend/taskpane/bridge-client.js";

afterEach(() => vi.unstubAllGlobals());

function pane(handlers: Record<string, unknown> = {}) {
  const send = vi.fn();
  const client = new BridgeClient({ onState() {}, onHello() {}, onProviderStatus() {}, onProviderEvent() {},
    onError(message: string) { throw new Error(message); }, ...handlers } as never, "ws://localhost");
  (client as unknown as { ws: unknown }).ws = { readyState: WebSocket.OPEN, send, close: vi.fn() };
  (client as unknown as { sessionToken: string }).sessionToken = "token";
  const receive = (raw: string) => (client as unknown as { onMessage(raw: string): Promise<void> }).onMessage(raw);
  return { client, receive, sent: () => send.mock.calls.map(([raw]) => JSON.parse(raw as string)) };
}
const goToRequest = (args: unknown) => JSON.stringify({ type: "tool_request", tool: "__go_to", callId: "c1", requestId: "r1", timeoutMs: 15_000, args });

describe("a request from another file's chat to show a range", () => {
  it("shows it and answers, without asking permission or recording a change", async () => {
    const goTo = vi.fn(async () => undefined);
    const runTool = vi.fn();
    const onToolExecuted = vi.fn();
    const { receive, sent } = pane({ goTo, runTool, onToolExecuted });
    await receive(goToRequest({ sheet: "Sheet1", address: "A1:E8" }));
    expect(goTo).toHaveBeenCalledWith({ sheet: "Sheet1", address: "A1:E8" });
    expect(sent()).toEqual([{ type: "tool_result", callId: "c1", result: { ok: true, data: {} } }]);
    expect(runTool).not.toHaveBeenCalled();
    expect(onToolExecuted).not.toHaveBeenCalled();
  });

  it("shows a range on the open sheet when none is named", async () => {
    const goTo = vi.fn(async () => undefined);
    const { receive } = pane({ goTo });
    await receive(goToRequest({ address: "B2" }));
    expect(goTo).toHaveBeenCalledWith({ address: "B2" });
  });

  it("says why when the workbook cannot show it", async () => {
    const { receive, sent } = pane({ goTo: async () => { throw new Error('There is no sheet named "Nope" in this workbook.'); } });
    await receive(goToRequest({ sheet: "Nope", address: "A1" }));
    expect(sent()).toEqual([{ type: "tool_result", callId: "c1", result: { ok: false, error: { code: "office_error", message: 'There is no sheet named "Nope" in this workbook.' } } }]);
  });

  it("refuses a request that names no range, and a pane that cannot show ranges", async () => {
    const goTo = vi.fn();
    const withHandler = pane({ goTo });
    await withHandler.receive(goToRequest({ sheet: "Sheet1" }));
    await withHandler.receive(goToRequest({ address: 5 }));
    expect(goTo).not.toHaveBeenCalled();
    expect(withHandler.sent().map((m) => m.result.ok)).toEqual([false, false]);
    const without = pane();
    await without.receive(goToRequest({ address: "A1" }));
    expect(without.sent()[0].result).toMatchObject({ ok: false, error: { code: "unsupported" } });
  });
});

describe("asking the bridge to show a range in another file", () => {
  it("sends the file, sheet and range with the session token", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ ok: true }), { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);
    const { client } = pane();
    await client.goToFile({ file: "Book2.xlsx", sheet: "Sheet1", address: "A1:E8" });
    expect(fetchMock).toHaveBeenCalledOnce();
    const [url, options] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("/session/go-to");
    expect(options.method).toBe("POST");
    expect((options.headers as Record<string, string>).authorization).toBe("Bearer token");
    expect(JSON.parse(options.body as string)).toEqual({ file: "Book2.xlsx", sheet: "Sheet1", address: "A1:E8" });
  });

  it("fails with the bridge's reason", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ ok: false, error: '"Book2.xlsx" is not open in Excel with this add-in.' }), { status: 200 })));
    await expect(pane().client.goToFile({ file: "Book2.xlsx", address: "A1" })).rejects.toThrow('"Book2.xlsx" is not open in Excel with this add-in.');
  });

  it("fails when the bridge refuses or is not connected", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ error: "unauthorized" }), { status: 401 })));
    await expect(pane().client.goToFile({ file: "Book2.xlsx", address: "A1" })).rejects.toThrow("unauthorized");
    const offline = pane();
    (offline.client as unknown as { ws: unknown }).ws = null;
    await expect(offline.client.goToFile({ file: "Book2.xlsx", address: "A1" })).rejects.toThrow("not connected");
  });
});
