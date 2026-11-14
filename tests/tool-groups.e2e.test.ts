import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { SkillStore } from "../bridge/src/skills.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { StartSessionContext } from "../providers/src/types.js";
import { CONNECTED_FILE_TOOLS, WEB_TESTING_TOOLS } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

/** Counts the runtimes it opens and what each was told it may offer. */
class Recording extends MockAdapter {
  starts: StartSessionContext[] = [];
  closes: string[] = [];
  constructor() { super({ id: "claude" }); }
  override async startSession(ctx: StartSessionContext) { this.starts.push(ctx); return `runtime-${this.starts.length}`; }
  override async closeSession(id: string) { this.closes.push(id); }
}

let bridge: Bridge | null = null;
const panes: PaneClient[] = [];
let folder = "";
afterEach(async () => {
  for (const pane of panes.splice(0)) await pane.close();
  await bridge?.close(); bridge = null;
  if (folder) fs.rmSync(folder, { recursive: true, force: true });
});

async function start() {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "tool-groups-"));
  const port = await getFreePort();
  const adapter = new Recording();
  bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), {
    usageFile: null, savedReferencesFile: null, workspaceConfigPath: null,
    skillUploadDir: path.join(folder, "uploaded"), skills: new SkillStore([path.join(folder, "uploaded")], path.join(folder, "templates")),
    systemPromptFile: path.join(folder, "system-prompt.txt"), registry: new ProviderRegistry().register(adapter),
  });
  const { url } = await bridge.listen();
  return { adapter, wsUrl: url.replace("http", "ws") + "/ws" };
}

async function open(wsUrl: string, id: string, name: string) {
  const pane = await PaneClient.connect(wsUrl);
  panes.push(pane);
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: id });
  await pane.waitFor(m => m.type === "hello_ok");
  pane.onMessage = m => { if (m.type === "tool_request" && m.tool === "__workbook_name") pane.send({ type: "tool_result", callId: m.callId!, result: { ok: true, data: { name } } }); };
  pane.send({ type: "workbook_info", name });
  pane.send({ type: "set_provider", provider: "claude" });
  await pane.waitFor(m => m.type === "provider_status");
  return pane;
}

const settings = { model: null, effort: "medium" as const, chatSessionId: "chat", capabilities: { webSearch: false, playwright: false } };
const until = async (condition: () => boolean) => { for (let i = 0; i < 300 && !condition(); i++) await new Promise(resolve => setTimeout(resolve, 10)); };
const env = (ctx: StartSessionContext) => ctx.mcp.env;

describe("tools offered only when they apply", () => {
  it("offers the tools for other files only while another file is open, and opens the runtime again when that changes", async () => {
    const { adapter, wsUrl } = await start();
    const first = await open(wsUrl, "first-file", "Budget.xlsx");
    first.send({ type: "prepare", ...settings });
    await until(() => adapter.starts.length === 1);
    expect(env(adapter.starts[0]!).EXCEL_CONNECTED_FILES).toBeUndefined();

    const second = await open(wsUrl, "second-file", "Actuals.xlsx");
    await until(() => adapter.starts.length === 2);
    expect(env(adapter.starts[1]!).EXCEL_CONNECTED_FILES).toBe("1");
    expect(adapter.closes).toContain("runtime-1");

    // Nothing changes while the same files stay open, however often the pane says what it runs with.
    first.send({ type: "prepare", ...settings });
    await new Promise(resolve => setTimeout(resolve, 150));
    expect(adapter.starts).toHaveLength(2);

    await second.close();
    await until(() => adapter.starts.length === 3);
    expect(env(adapter.starts[2]!).EXCEL_CONNECTED_FILES).toBeUndefined();
  });

  it("names the groups once, so the protocol, the MCP server and the providers cannot disagree about them", () => {
    expect(CONNECTED_FILE_TOOLS).toEqual(["list_connected_files", "read_connected_file", "send_to_connected_file"]);
    expect(WEB_TESTING_TOOLS).toContain("record_test_results");
  });
});
