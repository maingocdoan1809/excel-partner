import { afterEach, describe, expect, it } from "vitest";
import http from "node:http";
import WebSocket from "ws";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;

afterEach(async () => {
  await bridge?.close();
  bridge = null;
});

async function startBridge() {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude" })) });
  await bridge.listen();
  return port;
}

/** `fetch` forbids overriding Host, so use http.request to exercise rebinding-style requests. */
function request(port: number, headers: Record<string, string>, pathname = "/health") {
  return new Promise<{ status: number; headers: http.IncomingHttpHeaders }>((resolve, reject) => {
    const req = http.request({ host: "127.0.0.1", port, path: pathname, headers }, (res) => {
      res.resume();
      res.on("end", () => resolve({ status: res.statusCode ?? 0, headers: res.headers }));
    });
    req.on("error", reject);
    req.end();
  });
}

function openSocket(port: number, headers: Record<string, string>) {
  return new Promise<"open" | number>((resolve) => {
    const ws = new WebSocket(`ws://127.0.0.1:${port}/ws`, { headers });
    ws.once("open", () => { ws.close(); resolve("open"); });
    ws.once("unexpected-response", (_req, res) => { res.resume(); resolve(res.statusCode ?? 0); });
    ws.once("error", () => resolve(0));
  });
}

describe("bridge origin and host checks", () => {
  it("serves the task pane's own origin and clients that send no Origin", async () => {
    const port = await startBridge();
    expect((await request(port, {})).status).toBe(200);
    expect((await request(port, { origin: `http://127.0.0.1:${port}` })).status).toBe(200);
    expect((await request(port, { origin: `http://localhost:${port}`, host: `localhost:${port}` })).status).toBe(200);
    const pane = await PaneClient.connect(`ws://127.0.0.1:${port}/ws`);
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef" });
    await pane.waitFor((m) => m.type === "hello_ok");
    await pane.close();
  });

  it("rejects HTTP requests from other web origins and never grants cross-origin read access", async () => {
    const port = await startBridge();
    const foreign = await request(port, { origin: "https://evil.example" });
    expect(foreign.status).toBe(403);
    expect(foreign.headers["access-control-allow-origin"]).toBeUndefined();
    expect((await request(port, { origin: "null" })).status).toBe(403);
    expect((await request(port, {})).headers["access-control-allow-origin"]).toBeUndefined();
  });

  it("refuses the WebSocket handshake from a web page, so no session token is issued", async () => {
    const port = await startBridge();
    expect(await openSocket(port, { origin: "https://evil.example" })).toBe(403);
    expect(await openSocket(port, { origin: "null" })).toBe(403);
    expect(await openSocket(port, { origin: `http://127.0.0.1:${port}` })).toBe("open");
    expect(await openSocket(port, {})).toBe("open");
  });

  it("refuses a rebound hostname even when Origin matches Host", async () => {
    const port = await startBridge();
    const rebound = await request(port, { host: `evil.example:${port}`, origin: `http://evil.example:${port}` });
    expect(rebound.status).toBe(403);
    expect((await request(port, { host: `evil.example:${port}` })).status).toBe(403);
    expect(await openSocket(port, { host: `evil.example:${port}`, origin: `http://evil.example:${port}` })).toBe(403);
  });

  it("accepts extra hostnames configured in BRIDGE_ALLOWED_HOSTS", async () => {
    process.env.BRIDGE_ALLOWED_HOSTS = "wsl-box, 172.20.0.5";
    try {
      const port = await startBridge();
      expect((await request(port, { host: `wsl-box:${port}`, origin: `http://wsl-box:${port}` })).status).toBe(200);
      expect((await request(port, { host: `172.20.0.5:${port}` })).status).toBe(200);
      expect((await request(port, { host: `other:${port}` })).status).toBe(403);
    } finally { delete process.env.BRIDGE_ALLOWED_HOSTS; }
  });
});
