import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | undefined, pane: PaneClient | undefined, directory: string | undefined;
afterEach(async () => { await pane?.close(); await bridge?.close(); pane = undefined; bridge = undefined;
  vi.unstubAllEnvs(); if (directory) fs.rmSync(directory, { recursive: true, force: true }); directory = undefined; });
async function fixture() {
  directory = fs.mkdtempSync(path.join(os.tmpdir(), "system-controls-"));
  vi.stubEnv("OFFICE_PARTNER_DATA_DIR", directory); vi.stubEnv("OFFICE_PARTNER_CONTROL_TOKEN", "test-control-token");
  const port = await getFreePort(), shutdown = vi.fn();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, configFile: path.join(directory, "config.json"), logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, systemPromptFile: null, savedReferencesFile: null, updater: null, onShutdown: shutdown,
    registry: new ProviderRegistry().register(new MockAdapter({ id: "copilot", delayMs: 1000 })) });
  const { url } = await bridge.listen();
  pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef" });
  const hello = await pane.waitFor(message => message.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("Expected hello");
  const headers = { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" };
  const action = (name: string) => fetch(`${url}/session/system`, { method: "POST", headers, body: JSON.stringify({ action: name, provider: "copilot" }) });
  const control = (name: string, token = "test-control-token") => fetch(`${url}/admin/${name}`, { method: "POST", headers: { "x-office-partner-control": token } });
  return { url, headers, action, control, shutdown };
}
describe("authenticated recovery service controls", () => {
  it("requires a connected pane token for system metadata and actions", async () => {
    const w = await fixture();
    expect((await fetch(`${w.url}/session/system`)).status).toBe(401);
    expect((await fetch(`${w.url}/session/system`, { method: "POST", body: '{"action":"debug-on"}' })).status).toBe(401);
    const metadata = await (await fetch(`${w.url}/session/system`, { headers: w.headers })).json();
    expect(metadata).toMatchObject({ providers: expect.any(Array), cached: [] });
    expect(JSON.stringify(metadata)).not.toContain(directory!);
    expect((await w.action("debug-on")).status).toBe(200);
    expect(JSON.parse(fs.readFileSync(path.join(directory!, "debug.json"), "utf8")).expires).toBeGreaterThan(Date.now());
    expect((await w.action("remove")).status).toBe(200);
    expect(JSON.parse(fs.readFileSync(path.join(directory!, "runtime-state.json"), "utf8")).disabled).toContain("copilot");
  });

  it("authenticates maintenance, blocks new prompts, resumes, then shuts down through its callback", async () => {
    const w = await fixture();
    expect((await w.control("maintenance", "wrong")).status).toBe(401);
    expect((await w.control("maintenance")).status).toBe(200);
    expect((await w.action("debug-on")).status).toBe(503);
    pane!.send({ type: "set_provider", provider: "copilot" });
    await pane!.waitFor(message => message.type === "provider_status");
    pane!.send({ type: "prompt", requestId: "during-update", prompt: "hello", selection: null });
    expect(await pane!.waitFor(message => message.type === "turn_finished" && message.requestId === "during-update")).toMatchObject({ reason: "error" });
    expect((await w.control("resume")).status).toBe(200);
    expect((await w.action("debug-off")).status).toBe(200);
    expect((await w.control("shutdown")).status).toBe(200);
    await new Promise(resolve => setImmediate(resolve)); expect(w.shutdown).toHaveBeenCalledOnce();
  });

  it("rejects maintenance, shutdown and system changes during a response without entering maintenance", async () => {
    const w = await fixture();
    pane!.send({ type: "set_provider", provider: "copilot" }); await pane!.waitFor(message => message.type === "provider_status");
    pane!.send({ type: "prompt", requestId: "busy", prompt: "hello", selection: null });
    await pane!.waitFor(message => message.type === "turn_started");
    expect((await w.control("maintenance")).status).toBe(409);
    expect((await w.control("shutdown")).status).toBe(409); expect(w.shutdown).not.toHaveBeenCalled();
    expect((await w.action("debug-on")).status).toBe(409);
    expect(await (await fetch(`${w.url}/health`)).json()).toMatchObject({ maintenance: false });
  });
});
