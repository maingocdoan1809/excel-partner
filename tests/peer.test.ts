import fs from "node:fs";
import path from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { createPeerServer, httpPeerClient } from "../excel_mcp/src/peer.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;
let client: Client | null = null;
afterEach(async () => { await client?.close(); client = null; await bridge?.close(); bridge = null; });

async function start() {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", reply: "Totals written to D2:D40." })) });
  const { url } = await bridge.listen();
  const tokenFile = path.join(path.dirname(config.configFile), "peer-token.txt");
  return { url, wsUrl: `${url.replace("http", "ws")}/ws`, tokenFile };
}

/** What a terminal agent sees: the MCP server over a transport, talking to the bridge over HTTP. */
async function terminalAgent(url: string, tokenFile: string): Promise<Client> {
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  await createPeerServer(httpPeerClient({ bridgeUrl: url, token: () => fs.readFileSync(tokenFile, "utf8").trim() })).connect(serverSide);
  client = new Client({ name: "claude-code", version: "1" });
  await client.connect(clientSide);
  return client;
}

async function call(agent: Client, name: string, args: Record<string, unknown> = {}) {
  const result = await agent.callTool({ name, arguments: args });
  return { isError: result.isError === true, body: JSON.parse((result.content as Array<{ text: string }>)[0]!.text) };
}

/** A task pane for "Actuals.xlsx" doing what the real one does: answers the approval, then runs the delegated request as its own turn. */
async function openPane(wsUrl: string, approve = true) {
  const pane = await PaneClient.connect(wsUrl);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "actuals-1" });
  await pane.waitFor((m) => m.type === "hello_ok");
  pane.send({ type: "workbook_info", name: "Actuals.xlsx" });
  pane.send({ type: "set_provider", provider: "claude" });
  await pane.waitFor((m) => m.type === "provider_status");
  const seen: { approvals: unknown[]; delegated: unknown[] } = { approvals: [], delegated: [] };
  pane.onMessage = (m) => {
    if (m.type === "tool_request" && m.tool === "__workbook_name") pane.send({ type: "tool_result", callId: m.callId, result: { ok: true, data: { name: "Actuals.xlsx" } } });
    if (m.type === "tool_request" && m.tool === "__authorize_connected_file") {
      seen.approvals.push(m.args);
      pane.send({ type: "tool_result", callId: m.callId, result: { ok: true, data: { approved: approve } } });
    }
    if (m.type === "delegated_prompt") {
      seen.delegated.push({ from: m.from, text: m.text });
      pane.send({ type: "prompt", requestId: "turn-1", prompt: m.text, selection: null, model: null, effort: "medium", delegationId: m.delegationId });
    }
  };
  return { pane, seen };
}

describe("a terminal agent asking the assistant of an open file", () => {
  it("offers two native tools with schemas", async () => {
    const { url, tokenFile } = await start();
    const agent = await terminalAgent(url, tokenFile);
    const tools = (await agent.listTools()).tools;
    expect(tools.map(tool => tool.name)).toEqual(["list_office_files", "ask_office_file"]);
    expect(tools[1]!.inputSchema).toMatchObject({ required: ["file", "message"] });
  });

  it("lists the open files and runs the request as a turn in that file's chat", async () => {
    const { url, wsUrl, tokenFile } = await start();
    const { pane, seen } = await openPane(wsUrl);
    const agent = await terminalAgent(url, tokenFile);

    const files = await call(agent, "list_office_files");
    expect(files.body).toEqual([expect.objectContaining({ name: "Actuals.xlsx", host: "excel", ready: true, busy: false })]);

    const reply = await call(agent, "ask_office_file", { file: files.body[0].id, message: "Fill column D with totals" });
    expect(reply).toMatchObject({ isError: false, body: { file: { name: "Actuals.xlsx" }, outcome: "completed", reply: "Totals written to D2:D40." } });
    // The person at that window was asked first, told who is asking and shown the exact text.
    expect(seen.approvals).toEqual([expect.objectContaining({ action: "message", from: "Claude Code (terminal)", text: "Fill column D with totals" })]);
    expect(seen.delegated).toEqual([{ from: { id: "terminal", name: "Claude Code (terminal)" }, text: "Fill column D with totals" }]);
    await pane.waitFor((m) => m.type === "turn_finished" && m.requestId === "turn-1");
    await pane.close();
  });

  it("does nothing in the file when the person declines", async () => {
    const { url, wsUrl, tokenFile } = await start();
    const { pane, seen } = await openPane(wsUrl, false);
    const agent = await terminalAgent(url, tokenFile);
    const [file] = (await call(agent, "list_office_files")).body;
    const reply = await call(agent, "ask_office_file", { file: file.id, message: "Delete everything" });
    expect(reply).toMatchObject({ isError: true, body: { code: "cancelled" } });
    expect(seen.delegated).toEqual([]);
    await pane.close();
  });

  it("says so when the file is unknown or nothing is open", async () => {
    const { url, tokenFile } = await start();
    const agent = await terminalAgent(url, tokenFile);
    expect((await call(agent, "list_office_files")).body).toEqual([]);
    expect(await call(agent, "ask_office_file", { file: "f99", message: "Hello" })).toMatchObject({ isError: true, body: { code: "invalid_arguments" } });
    expect(await call(agent, "ask_office_file", { file: "f99", message: "  " })).toMatchObject({ isError: true, body: { code: "invalid_arguments" } });
  });

  it("refuses a caller without the bridge's peer secret", async () => {
    const { url } = await start();
    expect((await fetch(`${url}/peer/files`)).status).toBe(401);
    expect((await fetch(`${url}/peer/ask`, { method: "POST", headers: { "x-office-partner-peer": "guess" }, body: "{}" })).status).toBe(401);
  });

  it("explains itself when the bridge is not running", async () => {
    const agent = await terminalAgent(`http://127.0.0.1:${await getFreePort()}`, (() => { const file = path.join(fs.mkdtempSync(path.join(process.env.TEMP ?? ".", "peer-")), "token"); fs.writeFileSync(file, "x"); return file; })());
    expect(await call(agent, "list_office_files")).toMatchObject({ isError: true, body: { code: "disconnected", message: expect.stringContaining("not reachable") } });
  });
});
