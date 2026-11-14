import http from "node:http";
import { expect, it } from "vitest";
import { watchOpencodeStatus } from "../providers/src/opencode/status.js";

it("O02 forwards only this session's safe retry state, then clears it", async () => {
  const controller = new AbortController();
  const updates: unknown[] = [];
  let requests = 0;
  const server = http.createServer((req, res) => {
    expect(req.headers.authorization).toBe(`Basic ${Buffer.from("excel-partner:test-password").toString("base64")}`);
    expect(req.url).toBe("/session/status?directory=%2Fscratch");
    res.setHeader("content-type", "application/json");
    res.end(JSON.stringify({ other: { type: "retry", attempt: 99 }, mine: ++requests === 1
      ? { type: "retry", attempt: 2, message: "The inference service is overloaded. PRIVATE_RESPONSE", next: Date.now() + 1000 }
      : { type: "busy" } }));
  });
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  try {
    const address = server.address();
    if (!address || typeof address === "string") throw Error("No port");
    await watchOpencodeStatus({ port: address.port, password: "test-password", cwd: "/scratch", sessionId: "mine", signal: controller.signal,
      onStatus: status => { updates.push(status); if (updates.length === 2) controller.abort(); },
    });
    expect(updates).toEqual([{ state: "retrying", attempt: 2, overloaded: true }, { state: "waiting" }]);
    expect(JSON.stringify(updates)).not.toContain("PRIVATE");
  } finally { controller.abort(); server.closeAllConnections(); await new Promise<void>(resolve => server.close(() => resolve())); }
});

it("O02 tolerates older OpenCode versions without the status endpoint", async () => {
  const server = http.createServer((_req, res) => { res.writeHead(404); res.end(); });
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  try {
    const address = server.address();
    if (!address || typeof address === "string") throw Error("No port");
    await watchOpencodeStatus({ port: address.port, password: "test", cwd: "/scratch", sessionId: "mine", signal: new AbortController().signal,
      onStatus: () => { throw Error("Unexpected state"); },
    });
  } finally { server.closeAllConnections(); await new Promise<void>(resolve => server.close(() => resolve())); }
});
