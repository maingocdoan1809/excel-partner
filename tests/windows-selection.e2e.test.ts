import { describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { mergeWindowsSelection } from "../bridge/src/windows-selection.js";
import { SelectionSnapshotSchema, type SelectionSnapshot } from "../protocol/src/index.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput } from "../providers/src/types.js";
import { PaneClient, getFreePort } from "./helpers.js";

describe("native selection routing", () => {
  it("provides the same selected group to get_selection and the provider prompt", async () => {
    const baseline = SelectionSnapshotSchema.parse({ workbookName: "Flow.xlsx", sheet: "Flow", sheetId: "sheet-1",
      address: "A1", rowCount: 1, columnCount: 1, values: [[1]], headers: null, truncated: false });
    const shapes = [{ name: "Box A", nativeId: "2", nativeType: 1, connector: false },
      { name: "Arrow", nativeId: "3", nativeType: 9, connector: true }];
    let captured: SelectionSnapshot | null = null;
    class CaptureAdapter extends MockAdapter {
      override async *sendTurn(input: SendTurnInput) {
        captured = input.selection;
        yield* super.sendTurn(input);
      }
    }
    const port = await getFreePort();
    const bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), {
      usageFile: null, savedReferencesFile: null,
      registry: new ProviderRegistry().register(new CaptureAdapter()),
      selectionEnricher: async selection => mergeWindowsSelection(selection, { status: "ok", workbookName: "Flow.xlsx", sheet: "Flow", shapes }),
    });
    const { url } = await bridge.listen();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    try {
      pane.onMessage = message => {
        if (message.type === "tool_request") pane.send({ type: "tool_result", callId: message.callId, result: { ok: true, data: baseline } });
      };
      pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "selection-test" });
      const hello = await pane.waitFor(message => message.type === "hello_ok");
      if (hello.type !== "hello_ok") throw new Error("Missing session");
      const response = await fetch(`${url}/internal/tool`, { method: "POST",
        headers: { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" },
        body: JSON.stringify({ tool: "get_selection", args: {} }) });
      const tool = await response.json() as { ok: boolean; data: SelectionSnapshot };
      expect(tool).toMatchObject({ ok: true, data: { selectedShapes: shapes, selectedShapesStatus: "complete", values: [] } });
      pane.send({ type: "set_provider", provider: "claude" });
      await pane.waitFor(message => message.type === "provider_status");
      pane.send({ type: "prompt", requestId: "selected-group", prompt: "inspect these", selection: baseline });
      await pane.waitFor(message => message.type === "turn_finished");
      expect(captured).toEqual(tool.data);
    } finally { await pane.close(); await bridge.close(); }
  });
});
