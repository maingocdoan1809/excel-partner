import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ProfiledAdapter } from "../providers/src/profiled.js";
import { MockAdapter } from "../providers/src/mock.js";
import type { ProviderRuntime, SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderEvent } from "../protocol/src/index.js";

const store = createRequire(import.meta.url)(path.resolve("bin/runtime-store.cjs"));
const directories: string[] = [];
afterEach(() => { for (const dir of directories.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function fixture() {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "profiled-adapter-")); directories.push(dataDir);
  const profiles = ["old", "new"].map(name => {
    const command = path.join(dataDir, name + ".js"); fs.writeFileSync(command, name);
    return { id: store.digest(name), mode: "external", provider: "copilot", command, adapterRoot: path.resolve("."), fingerprint: store.fileDigest(command) };
  });
  store.changeState(dataDir, (state: any) => {
    for (const profile of profiles) state.profiles[profile.id] = profile;
    state.active.copilot = profiles[0]!.id; state.legacy.copilot = profiles[0]!.id;
  });
  const workers: Worker[] = [];
  class Worker extends MockAdapter {
    start = vi.fn(async (_ctx: StartSessionContext) => `worker-${workers.indexOf(this)}`);
    turns = vi.fn(); failTurn = false;
    constructor(readonly runtime: ProviderRuntime) { super({ id: "copilot" }); }
    override startSession(ctx: StartSessionContext) { return this.start(ctx); }
    override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
      this.turns(input);
      yield { type: "session_info", provider: "copilot", requestId: input.requestId, providerSessionId: "native-" + input.requestId };
      if (this.failTurn) throw new Error("protocol version mismatch after send");
      yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
    }
  }
  const env = { OFFICE_PARTNER_DATA_DIR: dataDir, COPILOT_AUTO_UPDATE: "true" };
  const factory = vi.fn((runtime: ProviderRuntime) => { const worker = new Worker(runtime); workers.push(worker); return worker; });
  const adapter = new ProfiledAdapter("copilot", { env, bridgeUrl: "", mcp: { command: "node", args: [] }, logger: { debug: vi.fn(), warn: vi.fn() } }, factory);
  const context: StartSessionContext = { workbookConnectionId: "workbook", sessionId: "chat", signal: new AbortController().signal,
    mcp: { bridgeUrl: "", sessionId: "chat", token: "fixture", command: "node", args: [], env: {} } };
  const input = (id: string, requestId: string): SendTurnInput => ({ providerSessionId: id, requestId, prompt: "hello", selection: null, signal: context.signal });
  const activateNext = () => store.changeState(dataDir, (state: any) => { state.previous.copilot = profiles[0]!.id; state.active.copilot = profiles[1]!.id; });
  return { dataDir, profiles, workers, env, factory, adapter, context, input, activateNext };
}
async function collect(events: AsyncIterable<ProviderEvent>) { const result = []; for await (const event of events) result.push(event); return result; }

describe("provider worker version isolation", () => {
  it("keeps running and saved chats on the recorded CLI while new chats select the next version", async () => {
    const w = fixture(), oldId = await w.adapter.startSession(w.context);
    await collect(w.adapter.sendTurn(w.input(oldId, "original")));
    w.activateNext();
    const newId = await w.adapter.startSession({ ...w.context, sessionId: "second" });
    await collect(w.adapter.sendTurn(w.input(oldId, "existing-worker")));
    expect(w.workers[0]!.turns).toHaveBeenCalledTimes(2);
    expect(w.workers[1]!.runtime.env?.COPILOT_BIN).toBe(w.profiles[1]!.command);
    const resumed = await w.adapter.startSession({ ...w.context, resumeSessionId: "native-original" });
    expect(resumed).toBe(oldId); expect(newId).not.toBe(oldId);
    expect(w.workers[0]!.start).toHaveBeenLastCalledWith(expect.objectContaining({ resumeSessionId: "native-original" }));
    expect(w.workers[0]!.runtime.env?.COPILOT_AUTO_UPDATE).toBe("false");
    expect(w.env.COPILOT_AUTO_UPDATE).toBe("true");
  });

  it("falls back only on initial transport failure and never replays a failed turn", async () => {
    const w = fixture(); await w.adapter.startSession(w.context); w.activateNext();
    await w.adapter.detect();
    w.workers[1]!.start.mockRejectedValueOnce(new Error("protocol version mismatch"));
    const id = await w.adapter.startSession(w.context);
    expect(id).toBe("worker-0");
    w.workers[0]!.failTurn = true;
    await expect(collect(w.adapter.sendTurn(w.input(id, "sent-once")))).rejects.toThrow("after send");
    expect(w.workers[0]!.turns).toHaveBeenCalledOnce(); expect(w.workers[1]!.turns).not.toHaveBeenCalled();
  });

  it("does not hide authentication errors or change the runtime of a resumed chat", async () => {
    const w = fixture(); await w.adapter.startSession(w.context); w.activateNext(); await w.adapter.detect();
    w.workers[1]!.start.mockRejectedValueOnce(new Error("401 authentication required"));
    await expect(w.adapter.startSession(w.context)).rejects.toThrow("authentication required");
    expect(w.workers[0]!.start).toHaveBeenCalledOnce();
    store.bindConversation(w.dataDir, "copilot", "saved-new", w.profiles[1]!.id);
    w.workers[1]!.start.mockRejectedValueOnce(new Error("protocol version mismatch"));
    await expect(w.adapter.startSession({ ...w.context, resumeSessionId: "saved-new" })).rejects.toThrow("protocol version mismatch");
    expect(w.workers[0]!.start).toHaveBeenCalledOnce();
  });

  it("disables new sessions while preserving the runtime required by a saved chat", async () => {
    const w = fixture(); store.bindConversation(w.dataDir, "copilot", "saved", w.profiles[0]!.id);
    store.changeState(w.dataDir, (state: any) => { state.disabled = ["copilot"]; });
    expect(await w.adapter.detect()).toMatchObject({ installed: false, status: "not_installed" });
    await expect(w.adapter.startSession(w.context)).rejects.toThrow("PROVIDER_DISABLED");
    await expect(w.adapter.startSession({ ...w.context, resumeSessionId: "saved" })).resolves.toBe("worker-0");
  });
});
