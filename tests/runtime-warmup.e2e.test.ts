import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { SkillStore } from "../bridge/src/skills.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderDetection, ProviderEvent } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

/** A provider whose runtime takes a while to open, so what is waiting for it can be seen. */
class SlowAdapter extends MockAdapter {
  starts: StartSessionContext[] = [];
  closes: string[] = [];
  detects = 0;
  releaseStart: (() => void) | null = null;
  holdStart = false;
  constructor() { super({ id: "claude" }); }
  override async detect(): Promise<ProviderDetection> {
    this.detects += 1;
    return super.detect();
  }
  override async startSession(context: StartSessionContext) {
    this.starts.push(context);
    if (this.holdStart) await new Promise<void>(resolve => { this.releaseStart = resolve; });
    return `runtime-${this.starts.length}`;
  }
  override async closeSession(id: string) { this.closes.push(id); }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    yield { type: "session_info", requestId: input.requestId, provider: "claude", providerSessionId: "native-session" };
    yield { type: "turn_started", requestId: input.requestId };
    yield { type: "message_completed", requestId: input.requestId, text: "done" };
    yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
  }
}

let bridge: Bridge | null = null;
let pane: PaneClient | null = null;
let folder = "";
afterEach(async () => {
  await pane?.close(); pane = null;
  await bridge?.close(); bridge = null;
  if (folder) fs.rmSync(folder, { recursive: true, force: true });
});

async function setup() {
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "warmup-"));
  const port = await getFreePort();
  const adapter = new SlowAdapter();
  bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), {
    usageFile: null, savedReferencesFile: null, workspaceConfigPath: null,
    skillUploadDir: path.join(folder, "uploaded"), skills: new SkillStore([path.join(folder, "uploaded")], path.join(folder, "templates")),
    systemPromptFile: path.join(folder, "system-prompt.txt"), registry: new ProviderRegistry().register(adapter),
  });
  const { url } = await bridge.listen();
  pane = await PaneClient.connect(url.replace("http", "ws") + "/ws");
  pane.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "test-workbook" });
  await pane.waitFor(message => message.type === "hello_ok");
  pane.send({ type: "set_provider", provider: "claude" });
  await pane.waitFor(message => message.type === "provider_status");
  return { adapter, pane };
}

const settings = { model: null, effort: "medium" as const, chatSessionId: "chat-1", capabilities: { webSearch: false, playwright: false } };
const until = async (condition: () => boolean) => { for (let i = 0; i < 200 && !condition(); i++) await new Promise(resolve => setTimeout(resolve, 10)); };
const prompt = (requestId: string) => ({ type: "prompt" as const, requestId, prompt: "hello", selection: null, chatSessionId: "chat-1", model: null, effort: "medium" as const,
  capabilities: { webSearch: false, playwright: false } });

describe("opening the provider's runtime ahead of the first message", () => {
  it("opens it on `prepare`, and the first prompt reuses it without checking the CLI again", async () => {
    const { adapter, pane } = await setup();
    pane.send({ type: "prepare", ...settings });
    await until(() => adapter.starts.length === 1);
    expect(adapter.starts).toHaveLength(1);

    const detectsBefore = adapter.detects;
    pane.send(prompt("r1"));
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "r1");
    // Like the pane, which sends the conversation's native id from the second message on.
    pane.send({ ...prompt("r2"), resumeSessionId: "native-session" });
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "r2");

    expect(adapter.starts).toHaveLength(1);
    expect(adapter.closes).toEqual([]);
    expect(adapter.detects).toBe(detectsBefore);
  });

  it("opens one runtime however often the same settings are sent", async () => {
    const { adapter, pane } = await setup();
    for (let i = 0; i < 4; i++) pane.send({ type: "prepare", ...settings });
    await until(() => adapter.starts.length === 1);
    await new Promise(resolve => setTimeout(resolve, 100));
    expect(adapter.starts).toHaveLength(1);
  });

  it("makes a prompt that arrives during the warm-up wait for it instead of opening a second runtime", async () => {
    const { adapter, pane } = await setup();
    adapter.holdStart = true;
    pane.send({ type: "prepare", ...settings });
    await until(() => adapter.starts.length === 1);
    pane.send(prompt("r1"));
    await pane.waitFor(message => message.type === "session_starting");
    adapter.releaseStart!();
    await pane.waitFor(message => message.type === "turn_finished" && message.requestId === "r1");
    expect(adapter.starts).toHaveLength(1);
  });

  it("keeps the runtime for as long as the connection lasts and closes it when the connection does", async () => {
    const { adapter, pane } = await setup();
    pane.send({ type: "prepare", ...settings });
    await until(() => adapter.starts.length === 1);
    await pane.close();
    await until(() => adapter.closes.length === 1);
    expect(adapter.closes).toEqual(["runtime-1"]);
  });
});
