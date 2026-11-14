import http from "node:http";
import net from "node:net";
import { spawn } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";
import WebSocket from "ws";
import { DEFAULT_PROVIDER } from "../protocol/src/index.js";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { GatewayServer } from "../gateway/src/server.js";
import { ConfigSchema } from "../gateway/src/config.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { isSameUserPeer, ownerIsCurrentUser, type PeerCheck } from "../shared/src/peer-identity.js";
import { PaneClient, getFreePort } from "./helpers.js";

const closers: Array<() => Promise<unknown>> = [];
afterEach(async () => { for (const close of closers.splice(0)) await close(); });

async function startBridge(peerCheck?: PeerCheck) {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  const bridge = new Bridge(config, { usageFile: null, peerCheck, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude" })) });
  closers.push(() => bridge.close());
  return (await bridge.listen()).url;
}

/** "refused" when the server drops the connection before answering. */
function get(url: string) {
  return new Promise<number | "refused">(resolve => {
    http.get(url, res => { res.resume(); resolve(res.statusCode ?? 0); }).on("error", () => resolve("refused"));
  });
}
function openSocket(url: string) {
  return new Promise<"open" | "refused">(resolve => {
    const ws = new WebSocket(url);
    ws.once("open", () => { ws.close(); resolve("open"); });
    ws.once("error", () => resolve("refused"));
  });
}

describe("only processes of the same Windows user may connect", () => {
  it.runIf(process.platform === "win32")("identifies another process of this user as the same user, over IPv4 and IPv6 loopback", async () => {
    for (const host of ["127.0.0.1", "::1"]) {
      const verdicts: boolean[] = [];
      const server = net.createServer(socket => { verdicts.push(ownerIsCurrentUser(socket)); socket.end(); });
      await new Promise<void>(resolve => server.listen(0, host, resolve));
      const { port } = server.address() as net.AddressInfo;
      await new Promise(resolve => spawn(process.execPath, ["-e", `require("net").connect(${port}, ${JSON.stringify(host)}).resume()`]).on("exit", resolve));
      server.close();
      expect(verdicts).toEqual([true]);
    }
  });

  it("decides once per connection and treats a failing check as foreign", () => {
    const socket = new net.Socket();
    let calls = 0;
    expect(isSameUserPeer(socket, () => { calls++; return true; })).toBe(true);
    expect(isSameUserPeer(socket, () => { calls++; return false; })).toBe(true);
    expect(calls).toBe(1);
    expect(isSameUserPeer(new net.Socket(), () => { throw new Error("lookup failed"); })).toBe(false);
  });

  it("serves the bridge to this user's processes and drops HTTP and WebSocket connections from another user", async () => {
    const own = await startBridge();
    expect(await get(`${own}/health`)).toBe(200);
    expect(await openSocket(`${own.replace("http", "ws")}/ws`)).toBe("open");
    const foreign = await startBridge(() => false);
    expect(await get(`${foreign}/health`)).toBe("refused");
    expect(await openSocket(`${foreign.replace("http", "ws")}/ws`)).toBe("refused");
  });

  it("never lists sessions or connection ids on the unauthenticated health check", async () => {
    const url = await startBridge();
    const pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    closers.push(() => pane.close());
    pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef" });
    await pane.waitFor(message => message.type === "hello_ok");
    const health = await (await fetch(`${url}/health`)).text();
    expect(JSON.parse(health)).toMatchObject({ status: "ok" });
    expect(health).not.toContain("0123456789abcdef");
    expect(health).not.toContain("sessions");
  });

  it("drops gateway connections from another user before the token is even read", async () => {
    const config = { ...ConfigSchema.parse({}), port: 0 };
    const server = new GatewayServer({ config, token: "t".repeat(32), tls: null, relayEntry: import.meta.filename, peerCheck: () => false,
      registry: new ProviderRegistry().register(new MockAdapter({ id: DEFAULT_PROVIDER })) });
    closers.push(() => server.close());
    expect(await get(`${await server.listen()}/health`)).toBe("refused");
  });
});
