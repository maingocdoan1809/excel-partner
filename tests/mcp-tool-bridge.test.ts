import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { connectMcpTools, type McpToolBridge } from "../providers/src/mcp-tool-bridge.js";

const fixture = path.resolve("tests/support/echo-mcp-server.mjs");
let bridge: McpToolBridge | undefined;
afterEach(async () => { await bridge?.close().catch(() => {}); bridge = undefined; });
const open = async () => (bridge = await connectMcpTools({ command: process.execPath, args: [fixture], env: { FIXTURE_VALUE: "from-env" } }));

describe("MCP tool bridge", () => {
  it("lists the tools the server describes, with their descriptions and input schemas", async () => {
    const { tools } = await open();
    expect(tools.map(tool => tool.name)).toEqual(["echo", "fail", "hang"]);
    expect(tools[0]).toEqual({ name: "echo", description: "Echo the text back",
      inputSchema: { type: "object", properties: { text: { type: "string" } }, required: ["text"] } });
  });
  it("runs a call in the server process, which gets the environment it was given", async () => {
    const opened = await open();
    expect(await opened.call("echo", { text: "hi" })).toEqual({ content: [{ type: "text", text: "echo:hi:from-env" }] });
  });
  it("returns a tool's own failure as a result, so the model can read it", async () => {
    const opened = await open();
    expect(await opened.call("fail", {})).toMatchObject({ isError: true, content: [{ type: "text", text: "it failed" }] });
  });
  it("stops waiting for a call when it is cancelled, and refuses calls once closed", async () => {
    const opened = await open();
    const stop = new AbortController();
    const pending = opened.call("hang", {}, stop.signal);
    setTimeout(() => stop.abort(), 50);
    await expect(pending).rejects.toThrow();
    await opened.close();
    await expect(opened.call("echo", { text: "late" })).rejects.toThrow();
  });
  it("fails to connect, and leaves nothing running, when the server cannot start", async () => {
    await expect(connectMcpTools({ command: process.execPath, args: [path.resolve("tests/support/does-not-exist.mjs")], env: {} })).rejects.toThrow();
  });
});
