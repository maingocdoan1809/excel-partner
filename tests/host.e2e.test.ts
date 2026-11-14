import { describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import type { PromptSelectionSnapshot } from "../protocol/src/index.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { McpEndpoint, SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import { PaneClient, getFreePort } from "./helpers.js";

const excelSnapshot = { workbookName: "Book1.xlsx", sheet: "Sheet1", sheetId: "1", address: "A1", rowCount: 1, columnCount: 1, values: [[1]], headers: null, truncated: false };
const wordSnapshot = { host: "word", text: "Hello", cursor: false, paragraphs: 1, truncated: false };

async function harness() {
  const seen: { mcp: McpEndpoint[]; selections: Array<PromptSelectionSnapshot | null> } = { mcp: [], selections: [] };
  let enriched = 0;
  class CaptureAdapter extends MockAdapter {
    override async startSession(ctx: StartSessionContext) {
      seen.mcp.push(ctx.mcp);
      return super.startSession(ctx);
    }
    override async *sendTurn(input: SendTurnInput) {
      seen.selections.push(input.selection);
      yield* super.sendTurn(input);
    }
  }
  const port = await getFreePort();
  const bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error", allowOfficeJs: true }), {
    usageFile: null, savedReferencesFile: null,
    registry: new ProviderRegistry().register(new CaptureAdapter()),
    selectionEnricher: async (selection) => { enriched++; return selection; },
  });
  const { url } = await bridge.listen();
  const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  return { seen, bridge, pane, enriched: () => enriched };
}

async function hello(pane: PaneClient, connectionId: string, clientInfo?: Record<string, unknown>) {
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: connectionId, ...(clientInfo ? { clientInfo: clientInfo as never } : {}) });
  return pane.waitFor((message) => message.type === "hello_ok" || message.type === "hello_error");
}

async function prompt(pane: PaneClient, requestId: string, selection: unknown) {
  pane.send({ type: "set_provider", provider: "claude" });
  await pane.waitFor((message) => message.type === "provider_status");
  pane.send({ type: "prompt", requestId, prompt: "go", selection } as never);
  return pane.waitFor((message) => (message.type === "turn_finished" || message.type === "error") && message.requestId === requestId);
}

describe("host through the bridge", () => {
  it.each(["excel", "word", "powerpoint"] as const)("passes OFFICE_HOST=%s to the MCP process without a global tool list", async (host) => {
    const { seen, bridge, pane } = await harness();
    try {
      expect((await hello(pane, `host-${host}-0001`, { host })).type).toBe("hello_ok");
      expect((await prompt(pane, "r1", null)).type).toBe("turn_finished");
      expect(seen.mcp).toHaveLength(1);
      expect(seen.mcp[0]!.env.OFFICE_HOST).toBe(host);
      expect(seen.mcp[0]!.toolNames).toBeUndefined();
    } finally { await pane.close(); await bridge.close(); }
  });

  it("treats a pane with no clientInfo or no host as Excel", async () => {
    for (const info of [undefined, {}, { standalone: true }]) {
      const { seen, bridge, pane, enriched } = await harness();
      try {
        expect((await hello(pane, "legacy-pane-0001", info)).type).toBe("hello_ok");
        expect((await prompt(pane, "r1", excelSnapshot)).type).toBe("turn_finished");
        expect(seen.mcp[0]!.env.OFFICE_HOST).toBe("excel");
        expect(seen.selections[0]).toMatchObject({ host: "excel", workbookName: "Book1.xlsx" });
        // The machine's own read of the selection is made when the model asks for it (get_selection), not for every message.
        expect(enriched()).toBe(0);
      } finally { await pane.close(); await bridge.close(); }
    }
  });

  it("does not run the COM enrichment for a Word selection", async () => {
    const { seen, bridge, pane, enriched } = await harness();
    try {
      await hello(pane, "word-pane-00001", { host: "word" });
      expect((await prompt(pane, "r1", wordSnapshot)).type).toBe("turn_finished");
      expect(seen.selections[0]).toMatchObject({ host: "word", text: "Hello" });
      expect(enriched()).toBe(0);
    } finally { await pane.close(); await bridge.close(); }
  });

  it("refuses a selection from another host and one that is not a snapshot", async () => {
    const { seen, bridge, pane } = await harness();
    try {
      await hello(pane, "word-pane-00002", { host: "word" });
      const crossHost = await prompt(pane, "r1", excelSnapshot);
      expect(crossHost).toMatchObject({ type: "error", code: "bad_message" });
      const malformed = await prompt(pane, "r2", { host: "word", text: 1 });
      expect(malformed).toMatchObject({ type: "error", code: "bad_message" });
      expect(seen.selections).toHaveLength(0);
    } finally { await pane.close(); await bridge.close(); }
  });

  it("rejects a hello that names an unknown host", async () => {
    const { bridge, pane } = await harness();
    try {
      expect(await hello(pane, "bad-host-000001", { host: "outlook" })).toMatchObject({ type: "hello_error", code: "invalid_hello" });
    } finally { await pane.close(); await bridge.close(); }
  });

  it("routes Word tools to a Word pane and refuses Excel and PowerPoint tools there", async () => {
    const { url, bridge, pane } = await harnessWithUrl();
    try {
      const requests: Array<{ tool: string; args: unknown }> = [];
      pane.onMessage = (message) => {
        if (message.type === "tool_request") {
          requests.push({ tool: message.tool, args: message.args });
          pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: { summary: "done" } } });
        }
      };
      const hello = await hello_(pane, "word-routing-001", { host: "word" });
      const call = async (tool: string, args: unknown) => (await fetch(`${url}/internal/tool`, { method: "POST",
        headers: { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" }, body: JSON.stringify({ tool, args }) })).json() as Promise<{ ok: boolean; error?: { code: string }; data?: unknown }>;

      expect(await call("edit_doc_text", { old_text: "a", new_text: "b" })).toMatchObject({ ok: true });
      expect(await call("get_selection", {})).toMatchObject({ ok: true });
      expect(await call("execute_office_js", { code: "return 1;" })).toMatchObject({ ok: true });
      expect(requests.map((r) => r.tool)).toEqual(["edit_doc_text", "get_selection", "execute_office_js"]);
      // The bridge applies the schema's defaults before the pane sees the call.
      expect(requests[0]!.args).toMatchObject({ old_text: "a", new_text: "b", occurrence: 1, all: false });

      for (const tool of ["read_range", "write_range", "edit_slide_text", "edit_slide_xml"]) {
        expect(await call(tool, { address: "A1" }), tool).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
      }
      expect(await call("edit_doc_text", { old_text: "" })).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
      expect(requests).toHaveLength(3);
    } finally { await pane.close(); await bridge.close(); }
  });

  it("routes PowerPoint tools to a PowerPoint pane and refuses Word tools there", async () => {
    const { url, bridge, pane } = await harnessWithUrl();
    try {
      const requests: string[] = [];
      pane.onMessage = (message) => {
        if (message.type === "tool_request") { requests.push(message.tool); pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: {} } }); }
      };
      const hello = await hello_(pane, "ppt-routing-0001", { host: "powerpoint" });
      const call = async (tool: string, args: unknown) => (await fetch(`${url}/internal/tool`, { method: "POST",
        headers: { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" }, body: JSON.stringify({ tool, args }) })).json() as Promise<{ ok: boolean; error?: { code: string } }>;
      expect(await call("list_slide_shapes", { slide_id: "256" })).toMatchObject({ ok: true });
      expect(await call("edit_slide_xml", { slide_id: "256", code: "return 1;" })).toMatchObject({ ok: true });
      expect(await call("edit_doc_text", { old_text: "a", new_text: "b" })).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
      expect(await call("read_range", { address: "A1" })).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
      expect(requests).toEqual(["list_slide_shapes", "edit_slide_xml"]);
    } finally { await pane.close(); await bridge.close(); }
  });

  it("keeps Excel's own tools working for an Excel pane and refuses Word and PowerPoint tools", async () => {
    const { url, bridge, pane } = await harnessWithUrl();
    try {
      const requests: string[] = [];
      pane.onMessage = (message) => {
        if (message.type === "tool_request") { requests.push(message.tool); pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: {} } }); }
      };
      const hello = await hello_(pane, "excel-routing-001");
      const call = async (tool: string, args: unknown) => (await fetch(`${url}/internal/tool`, { method: "POST",
        headers: { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" }, body: JSON.stringify({ tool, args }) })).json() as Promise<{ ok: boolean; error?: { code: string } }>;
      expect(await call("get_sheets", {})).toMatchObject({ ok: true });
      expect(await call("edit_doc_text", { old_text: "a", new_text: "b" })).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
      expect(await call("list_slide_shapes", {})).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
      expect(requests).toEqual(["get_sheets"]);
    } finally { await pane.close(); await bridge.close(); }
  });
});

async function harnessWithUrl() {
  const port = await getFreePort();
  const bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error", allowOfficeJs: true }), {
    usageFile: null, savedReferencesFile: null, registry: new ProviderRegistry().register(new MockAdapter()),
  });
  const { url } = await bridge.listen();
  const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  return { url, bridge, pane };
}

async function hello_(pane: PaneClient, connectionId: string, clientInfo?: Record<string, unknown>) {
  const reply = await hello(pane, connectionId, clientInfo);
  if (reply.type !== "hello_ok") throw new Error("no handshake");
  return reply;
}
