import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { UpdateService } from "../bridge/src/updater.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

let folder = "";
let bridge: Bridge | null = null;
let pane: PaneClient | null = null;
beforeEach(() => { folder = fs.mkdtempSync(path.join(os.tmpdir(), "update-route-")); });
afterEach(async () => { await pane?.close(); pane = null; await bridge?.close(); bridge = null; fs.rmSync(folder, { recursive: true, force: true }); });

function packageFolder() {
  const root = path.join(folder, "global", "office-partner");
  fs.mkdirSync(path.join(root, "bin"), { recursive: true });
  for (const file of ["update.mjs", "update-runner.mjs"]) fs.copyFileSync(path.join("bin", file), path.join(root, "bin", file));
  fs.writeFileSync(path.join(root, "package.json"), JSON.stringify({ name: "office-partner", version: "0.1.19" }));
  return root;
}

async function connect(updater: UpdateService | null) {
  const port = await getFreePort();
  bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error", configFile: path.join(folder, "config.json") }), {
    usageFile: null, savedReferencesFile: null, workspaceConfigPath: null, skillUploadDir: path.join(folder, "skills"),
    systemPromptFile: path.join(folder, "system-prompt.txt"), registry: new ProviderRegistry().register(new MockAdapter({ id: "claude" })), updater,
  });
  const { url } = await bridge.listen();
  pane = await PaneClient.connect(url.replace("http", "ws") + "/ws");
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "update-workbook" });
  const hello = await pane.waitFor((message) => message.type === "hello_ok");
  if (hello.type !== "hello_ok") throw new Error("no handshake");
  return { url, headers: { authorization: `Bearer ${hello.sessionToken}` } };
}

describe("/session/update", () => {
  it("answers who is signed in with 401, and says it is off when there is no updater", async () => {
    const { url, headers } = await connect(null);
    expect((await fetch(url + "/session/update", { headers: { authorization: "Bearer nope" } })).status).toBe(401);
    expect((await fetch(url + "/session/update", { method: "POST", headers: { authorization: "Bearer nope" } })).status).toBe(401);
    expect((await fetch(url + "/session/update", { headers })).status).toBe(404);
  });

  it("reports the versions and asks the registry again on request", async () => {
    const fetchLatest = vi.fn(async () => "0.1.20");
    const { url, headers } = await connect(new UpdateService({ packageRoot: packageFolder(), dataDir: path.join(folder, "data"), platform: "win32", globalRoot: () => path.join(folder, "global"), fetchLatest }));
    const first = await (await fetch(url + "/session/update", { headers })).json() as Record<string, unknown>;
    expect(first).toMatchObject({ current: "0.1.19", latest: "0.1.20", updateAvailable: true, canUpdate: true, status: { state: "idle" } });
    await fetch(url + "/session/update", { headers });
    expect(fetchLatest).toHaveBeenCalledTimes(1);
    await fetch(url + "/session/update?refresh=1", { headers });
    expect(fetchLatest).toHaveBeenCalledTimes(2);
  });

  it("starts an update, then refuses a second one while the first runs", async () => {
    const launch = vi.fn(() => 321);
    const updater = new UpdateService({ packageRoot: packageFolder(), dataDir: path.join(folder, "data"), platform: "win32", globalRoot: () => path.join(folder, "global"), fetchLatest: async () => "0.1.20", launch, isPidAlive: () => true });
    const { url, headers } = await connect(updater);
    const started = await fetch(url + "/session/update", { method: "POST", headers });
    expect(started.status).toBe(202);
    expect(await started.json()).toMatchObject({ status: { state: "running", from: "0.1.19" } });
    expect(launch).toHaveBeenCalledOnce();
    fs.writeFileSync(path.join(folder, "data", "update-status.json"), JSON.stringify({ state: "running", pid: 321, startedAt: new Date().toISOString() }));
    const again = await fetch(url + "/session/update", { method: "POST", headers });
    expect(again.status).toBe(409);
    expect(await again.json()).toMatchObject({ code: "running" });
    expect(launch).toHaveBeenCalledOnce();
    fs.rmSync(path.dirname((launch.mock.calls[0] as unknown as [string])[0]), { recursive: true, force: true });
  });

  it("does not update while the assistant is working", async () => {
    const launch = vi.fn(() => 1);
    const updater = new UpdateService({ packageRoot: packageFolder(), dataDir: path.join(folder, "data"), platform: "win32", globalRoot: () => path.join(folder, "global"), fetchLatest: async () => "0.1.20", launch });
    const { url, headers } = await connect(updater);
    const sessions = (bridge as unknown as { sessions: { list(): Array<{ activeRequestId: string | null }> } }).sessions.list();
    sessions[0]!.activeRequestId = "turn-1";
    const refused = await fetch(url + "/session/update", { method: "POST", headers });
    expect(refused.status).toBe(409);
    expect(await refused.json()).toMatchObject({ code: "busy", error: expect.stringContaining("still working") });
    expect(launch).not.toHaveBeenCalled();
    sessions[0]!.activeRequestId = null;
    expect((await fetch(url + "/session/update", { method: "POST", headers })).status).toBe(202);
    fs.rmSync(path.dirname((launch.mock.calls[0] as unknown as [string])[0]), { recursive: true, force: true });
  });

  it("refuses with 400 when this copy cannot replace itself", async () => {
    const launch = vi.fn(() => 1);
    const updater = new UpdateService({ packageRoot: packageFolder(), dataDir: path.join(folder, "data"), platform: "win32", globalRoot: () => path.join(folder, "elsewhere"), fetchLatest: async () => "0.1.20", launch });
    const { url, headers } = await connect(updater);
    const response = await fetch(url + "/session/update", { method: "POST", headers });
    expect(response.status).toBe(400);
    expect(await response.json()).toMatchObject({ code: "unsupported" });
    expect(launch).not.toHaveBeenCalled();
  });
});
