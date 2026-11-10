import path from "node:path";
import { spawnSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { exposedToolNamesForHost, type HostId } from "../protocol/src/index.js";

const root = path.join(__dirname, "..");
const tsx = path.join(root, "node_modules", ".bin", process.platform === "win32" ? "tsx.cmd" : "tsx");
const cli = path.join(root, "excel_mcp", "src", "cli.ts");
const base = { EXCEL_BRIDGE_URL: "http://127.0.0.1:1", EXCEL_SESSION_ID: "s", EXCEL_SESSION_TOKEN: "t", EXCEL_ALLOW_OFFICE_JS: "1" };

async function listed(env: Record<string, string>) {
  const transport = new StdioClientTransport({ command: tsx, args: [cli], env: { ...process.env as Record<string, string>, ...base, ...env }, stderr: "pipe" });
  const client = new Client({ name: "test", version: "1" });
  await client.connect(transport);
  try {
    return (await client.listTools()).tools.map((tool) => tool.name).sort();
  } finally { await client.close(); }
}

describe("excel-mcp process per OFFICE_HOST", () => {
  it.each([[undefined, "excel"], ["excel", "excel"], ["word", "word"], ["powerpoint", "powerpoint"]] as Array<[string | undefined, HostId]>)("OFFICE_HOST=%s lists the %s tools", async (value, host) => {
    const tools = await listed(value ? { OFFICE_HOST: value } : {});
    expect(tools).toEqual(exposedToolNamesForHost(host, true).sort());
    if (host === "excel") expect(tools).toContain("read_range");
    else expect(tools).not.toContain("read_range");
  }, 30_000);

  it("refuses to start for an unknown host", () => {
    const run = spawnSync(process.execPath, [path.join(root, "node_modules", "tsx", "dist", "cli.mjs"), cli], { env: { ...process.env, ...base, OFFICE_HOST: "outlook" }, encoding: "utf8", timeout: 30_000 });
    expect(run.error).toBeUndefined();
    expect(run.status).toBe(2);
    expect(run.stderr).toMatch(/OFFICE_HOST/);
  }, 40_000);
});
