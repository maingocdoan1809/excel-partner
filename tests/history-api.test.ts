import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { MockAdapter } from "../providers/src/mock.js";
import { PaneClient } from "./helpers.js";
let bridge: Bridge, pane: PaneClient, dir: string;
afterEach(async () => { await pane?.close(); await bridge?.close(); if (dir) fs.rmSync(dir, { recursive: true, force: true }); });
async function fixture(host = "excel") {
 dir = fs.mkdtempSync(path.join(os.tmpdir(), "history-api-"));
 const config = loadConfig({ host: "127.0.0.1", port: 0, tls: null, configFile: path.join(dir, "config.json"), logLevel: "error" });
 bridge = new Bridge(config, { registry: new ProviderRegistry().register(new MockAdapter()), usageFile: null, systemPromptFile: null, savedReferencesFile: null, updater: null });
 const { url } = await bridge.listen(); pane = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
 pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "0123456789abcdef", clientInfo: { host } } as never);
 const hello = await pane.waitFor(m => m.type === "hello_ok"); if (hello.type !== "hello_ok") throw new Error("hello");
 const headers = { authorization: `Bearer ${hello.sessionToken}`, "content-type": "application/json" };
 const call = (body?: unknown, query = "") => fetch(`${url}/session/history${query}`, { headers, ...(body ? { method: "POST", body: JSON.stringify(body) } : {}) });
 const session = { id: "chat", provider: "copilot", host: "word", title: "test", startedAt: 1, updatedAt: 2, messages: [{ kind: "user", text: "hello" }], changes: [] };
 return { url, headers, call, session };
}
describe("authenticated history API", () => {
 it("H19 leaves corrupt storage intact while the rest of the bridge stays healthy", async () => {
   const w = await fixture(); fs.writeFileSync(path.join(dir, "history.db"), "broken");
   expect((await w.call()).status).toBe(503); expect(fs.readFileSync(path.join(dir, "history.db"), "utf8")).toBe("broken");
   expect((await fetch(`${w.url}/health`)).status).toBe(200);
 });
 it("H11 validates the whole migration batch before storing anything", async () => {
   const w = await fixture();
   expect((await w.call({ action: "import", sessions: [{ ...w.session, host: "excel" }, { id: "invalid" }] })).status).toBe(400);
   expect(await (await w.call()).json()).toMatchObject({ total: 0 });
 });

 it("H08/H09 derives host from the connected session and rejects unauthorized access", async () => {
  const w = await fixture(); expect((await fetch(`${w.url}/session/history`)).status).toBe(401);
  expect((await fetch(`${w.url}/session/history`, { headers: { authorization: "Bearer invalid-token" } })).status).toBe(401);
  expect((await w.call({ action: "save", session: w.session, baseRevision: 0, generation: 0, operationId: "first" })).status).toBe(200);
  const record = await (await w.call(undefined, "?view=chat&id=chat&host=word")).json();
  expect(record.session.host).toBe("excel");
  await pane.close(); expect((await w.call()).status).toBe(401);
 });
 it("H05/H07/H20 reports conflicts, exports and acknowledges scoped deletion", async () => {
  const w = await fixture(); const body = { action: "save", session: w.session, baseRevision: 0, generation: 0, operationId: "first" };
  await w.call(body); expect((await w.call({ ...body, operationId: "stale" })).status).toBe(409);
  expect(await (await w.call(undefined, "?view=export")).json()).toMatchObject({ sessions: [{ id: "chat" }] });
  expect((await fetch(`${w.url}/session/history?id=chat`, { method: "DELETE", headers: w.headers })).status).toBe(200);
  expect((await w.call(body)).status).toBe(409);
 });
});
