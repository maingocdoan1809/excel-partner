import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { toolCatalogChars } from "../shared/src/catalog.js";
import { GET_TOOL_DEFINITION } from "../shared/src/lazy-tools.js";
import { PaneClient, getFreePort } from "./helpers.js";

const GROUPS = { playwright: false, webSearch: false, connectedFiles: false };

/** One turn through a bridge whose own `toolSchemas` setting is `bridgeSetting`, from a pane that sends `capabilities`. */
async function turn(bridgeSetting: "full" | "lazy", capabilities: { webSearch: boolean; playwright: boolean; lazySchemas?: boolean }) {
  const folder = fs.mkdtempSync(path.join(os.tmpdir(), "lazy-pane-"));
  const file = path.join(folder, "usage.jsonl");
  const port = await getFreePort();
  const config = { ...loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), toolSchemas: bridgeSetting };
  const bridge = new Bridge(config, { usageFile: file, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", toolCalls: [
    { tool: GET_TOOL_DEFINITION, args: { names: ["set_cells"] } }] })) });
  const { url } = await bridge.listen();
  const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  try {
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: `lazy-pane-${bridgeSetting}` });
    const hello = await pane.waitFor(m => m.type === "hello_ok");
    pane.send({ type: "set_provider", provider: "claude" });
    await pane.waitFor(m => m.type === "provider_status");
    pane.send({ type: "prompt", requestId: "r1", chatSessionId: "chat", prompt: "describe", selection: null, capabilities });
    await pane.waitFor(m => m.type === "turn_finished" && m.requestId === "r1");
    const entry = fs.readFileSync(file, "utf8").split("\n").filter(Boolean).map(line => JSON.parse(line)).find(line => line.requestId === "r1");
    return { hello, lookup: pane.received.find(m => m.type === "tool_finished" && m.tool === GET_TOOL_DEFINITION), catalogChars: entry.catalogChars as number };
  } finally {
    await pane.close();
    await bridge.close();
    fs.rmSync(folder, { recursive: true, force: true });
  }
}

describe("the pane chooses how tools are offered", () => {
  it("tells the pane what the bridge does when the pane has not chosen", async () => {
    expect((await turn("full", { webSearch: false, playwright: false })).hello).toMatchObject({ toolSchemas: "full" });
    expect((await turn("lazy", { webSearch: false, playwright: false })).hello).toMatchObject({ toolSchemas: "lazy" });
  });

  it("follows the bridge's setting when the pane says nothing", async () => {
    const lazy = await turn("lazy", { webSearch: false, playwright: false });
    expect(lazy.lookup).toMatchObject({ ok: true });
    expect(lazy.catalogChars).toBe(toolCatalogChars("excel", true, { ...GROUPS, lazySchemas: true }));
    const full = await turn("full", { webSearch: false, playwright: false });
    expect(full.lookup).toMatchObject({ ok: false });
    expect(full.catalogChars).toBe(toolCatalogChars("excel", true, GROUPS));
  });

  it("lets the pane switch to the compact list on a bridge that offers full definitions", async () => {
    const result = await turn("full", { webSearch: false, playwright: false, lazySchemas: true });
    expect(result.lookup).toMatchObject({ ok: true });
    expect(result.catalogChars).toBe(toolCatalogChars("excel", true, { ...GROUPS, lazySchemas: true }));
  });

  it("lets the pane switch back to full definitions on a bridge set to the compact list", async () => {
    const result = await turn("lazy", { webSearch: false, playwright: false, lazySchemas: false });
    expect(result.lookup).toMatchObject({ ok: false });
    expect(result.catalogChars).toBe(toolCatalogChars("excel", true, GROUPS));
  });
});
