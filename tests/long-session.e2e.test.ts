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

const NATIVE = "native-session";

/** A provider that, like the real ones, keeps one conversation open and answers each message at once. */
class LongSessionAdapter extends MockAdapter {
  override readonly capabilities = { ...MockAdapter.capabilities, liveModel: true, liveEffort: true };
  starts: StartSessionContext[] = [];
  closes: string[] = [];
  turns: SendTurnInput[] = [];
  detects = 0;
  constructor() { super({ id: "claude" }); }
  override async detect(): Promise<ProviderDetection> { this.detects += 1; return super.detect(); }
  override async startSession(context: StartSessionContext) { this.starts.push(context); return `runtime-${this.starts.length}`; }
  override async closeSession(id: string) { this.closes.push(id); }
  override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
    this.turns.push(input);
    yield { type: "session_info", requestId: input.requestId, provider: "claude", providerSessionId: NATIVE };
    yield { type: "turn_started", requestId: input.requestId };
    yield { type: "message_completed", requestId: input.requestId, text: `answer ${this.turns.length}` };
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
  folder = fs.mkdtempSync(path.join(os.tmpdir(), "long-session-"));
  const port = await getFreePort();
  const adapter = new LongSessionAdapter();
  const systemPromptFile = path.join(folder, "system-prompt.txt");
  bridge = new Bridge(loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" }), {
    usageFile: null, savedReferencesFile: null, workspaceConfigPath: null,
    skillUploadDir: path.join(folder, "uploaded"), skills: new SkillStore([path.join(folder, "uploaded")], path.join(folder, "templates")),
    systemPromptFile, registry: new ProviderRegistry().register(adapter),
    selectionEnricher: async (selection: unknown) => selection,
  } as never);
  const { url } = await bridge.listen();
  const client = await PaneClient.connect(url.replace("http", "ws") + "/ws");
  pane = client;
  client.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "test-workbook" });
  await client.waitFor(message => message.type === "hello_ok");
  client.send({ type: "set_provider", provider: "claude" });
  await client.waitFor(message => message.type === "provider_status");
  return { adapter, pane: client, systemPromptFile };
}

/** What the pane does: after the first answer it names the conversation by the provider's own id. */
function conversation(client: PaneClient) {
  let chatSessionId = "chat-1";
  let count = 0;
  return async (extra: Record<string, unknown> = {}) => {
    const requestId = `r${++count}`;
    const selection = { workbookName: "Book1.xlsx", sheet: "Sheet1", sheetId: null, address: `A${count}:C${count}`, rowCount: 1, columnCount: 3, headers: ["a", "b", "c"],
      values: [[count, count + 1, count + 2]], truncated: false };
    const started = performance.now();
    client.send({ type: "prompt", requestId, prompt: `message ${count}`, selection, model: null, effort: "medium", chatSessionId,
      capabilities: { webSearch: false, playwright: false }, ...extra } as never);
    await client.waitFor(message => message.type === "turn_finished" && message.requestId === requestId);
    chatSessionId = `claude:${NATIVE}`;
    return performance.now() - started;
  };
}

describe("a long session", () => {
  it("keeps one runtime open for many messages and never checks the CLI again", async () => {
    const { adapter, pane } = await setup();
    const send = conversation(pane);
    await send();
    const detectsAfterFirst = adapter.detects;
    for (let i = 0; i < 39; i++) await send();

    expect(adapter.turns).toHaveLength(40);
    expect(adapter.starts).toHaveLength(1);
    expect(adapter.closes).toEqual([]);
    expect(adapter.detects).toBe(detectsAfterFirst);
    // Every message went to the one runtime, in order, and each carried its own selection.
    expect(new Set(adapter.turns.map(turn => turn.providerSessionId))).toEqual(new Set(["runtime-1"]));
    expect(adapter.turns[39]!.prompt).toContain("message 40");
    expect(adapter.turns[39]!.selection).toMatchObject({ address: "A40:C40" });
  });

  it("hands over a model or effort change inside the open runtime instead of restarting it", async () => {
    const { adapter, pane } = await setup();
    const send = conversation(pane);
    await send();
    await send({ model: "model-b", effort: "high" });
    await send({ model: "model-c", effort: "low" });
    await send({ model: null, effort: "medium" });

    expect(adapter.starts).toHaveLength(1);
    expect(adapter.closes).toEqual([]);
    expect(adapter.turns.map(turn => turn.model ?? null)).toEqual([null, "model-b", "model-c", null]);
    expect(adapter.turns.map(turn => turn.effort)).toEqual(["medium", "high", "low", "medium"]);
  });

  it("keeps the conversation when the instructions change, by opening the runtime again on the same conversation", async () => {
    const { adapter, pane, systemPromptFile } = await setup();
    const send = conversation(pane);
    for (let i = 0; i < 5; i++) await send();
    fs.writeFileSync(systemPromptFile, "New rules for the model.");
    for (let i = 0; i < 5; i++) await send();

    // One restart for the new instructions; the rest of the ten messages share the runtime that follows it.
    expect(adapter.starts).toHaveLength(2);
    expect(adapter.closes).toEqual(["runtime-1"]);
    expect(adapter.starts[1]!.resumeSessionId).toBe(NATIVE);
    expect(adapter.starts[1]!.systemPrompt).toContain("New rules for the model.");
    expect(adapter.turns.slice(0, 5).every(turn => turn.providerSessionId === "runtime-1")).toBe(true);
    expect(adapter.turns.slice(5).every(turn => turn.providerSessionId === "runtime-2")).toBe(true);
  });

  it("does not get slower as the session grows", async () => {
    const { pane } = await setup();
    const send = conversation(pane);
    await send();
    const times: number[] = [];
    for (let i = 0; i < 60; i++) times.push(await send());
    const median = (values: number[]) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)]!;
    const early = median(times.slice(0, 20));
    const late = median(times.slice(-20));
    // The bridge itself adds a few milliseconds per message; a leak or a growing list would show as a late median far above the early one.
    expect(late).toBeLessThan(Math.max(early * 3, 50));
  });
});
