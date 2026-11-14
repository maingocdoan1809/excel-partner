import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ProfiledAdapter } from "../providers/src/profiled.js";
import { MockAdapter } from "../providers/src/mock.js";
import { adapterClass } from "./support/adapter-class.js";
import type { ProviderRuntime, SendTurnInput, StartSessionContext } from "../providers/src/types.js";
import type { ProviderEvent } from "../protocol/src/index.js";

const store = createRequire(import.meta.url)(path.resolve("bin/runtime-store.cjs"));
const directories: string[] = [];
afterEach(() => { for (const dir of directories.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });

function fixture() {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "profiled-adapter-")); directories.push(dataDir);
  const installs = ["old", "new"].map(name => {
    const id = store.digest(name), command = path.join(dataDir, "runtimes", id, "cli.js");
    fs.mkdirSync(path.dirname(command), { recursive: true }); fs.writeFileSync(command, name);
    return { id, provider: "copilot", version: name, command, stateHome: path.join(dataDir, "provider-data", "copilot", "home"),
      fingerprint: store.fileDigest(command), platform: process.platform, arch: process.arch };
  });
  const select = (index: number) => store.changeState(dataDir, (state: any) => {
    for (const install of installs) state.profiles[install.id] = install;
    state.active.copilot = installs[index]!.id;
  });
  select(0);
  const workers: Worker[] = [];
  class Worker extends MockAdapter {
    start = vi.fn(async (_ctx: StartSessionContext) => `worker-${workers.indexOf(this)}`);
    turns = vi.fn();
    cancelled = vi.fn();
    constructor(readonly runtime: ProviderRuntime) { super({ id: "copilot" }); }
    override startSession(ctx: StartSessionContext) { return this.start(ctx); }
    override async cancelTurn(id: string) { this.cancelled(id); }
    override async *sendTurn(input: SendTurnInput): AsyncIterable<ProviderEvent> {
      this.turns(input);
      yield { type: "turn_finished", requestId: input.requestId, reason: "completed" };
    }
  }
  const factory = vi.fn((runtime: ProviderRuntime) => { const worker = new Worker(runtime); workers.push(worker); return worker; });
  const env = { OFFICE_PARTNER_DATA_DIR: dataDir, COPILOT_AUTO_UPDATE: "true" };
  const adapter = new ProfiledAdapter("copilot", { env, bridgeUrl: "", mcp: { command: "node", args: [] }, logger: { debug: vi.fn(), warn: vi.fn() } }, adapterClass(factory));
  const context: StartSessionContext = { workbookConnectionId: "workbook", sessionId: "chat", signal: new AbortController().signal,
    mcp: { bridgeUrl: "", sessionId: "chat", token: "fixture", command: "node", args: [], env: {} } };
  const input = (id: string, requestId: string): SendTurnInput => ({ providerSessionId: id, requestId, prompt: "hello", selection: null, signal: context.signal });
  return { dataDir, installs, workers, factory, adapter, context, input, select };
}
async function collect(events: AsyncIterable<ProviderEvent>) { const result = []; for await (const event of events) result.push(event); return result; }

describe("running a provider against the CLI the product installed", () => {
  it("gives the adapter the installed CLI, its own state folder and updates turned off", async () => {
    const w = fixture();
    await w.adapter.startSession(w.context);
    expect(w.workers[0]!.runtime.env).toMatchObject({
      COPILOT_BIN: w.installs[0]!.command, COPILOT_HOME: w.installs[0]!.stateHome, COPILOT_AUTO_UPDATE: "false",
    });
  });

  it("runs the system CLI when the product installed none", async () => {
    const w = fixture();
    store.changeState(w.dataDir, (state: any) => { state.active = {}; });
    await w.adapter.startSession(w.context);
    expect(w.workers[0]!.runtime.env).not.toHaveProperty("COPILOT_HOME");
    expect(w.workers[0]!.runtime.env).toMatchObject({ COPILOT_AUTO_UPDATE: "false" });
  });

  it("looks the install up on every call: a CLI switched while running is used by the next session, and each session stays with its own", async () => {
    const w = fixture();
    const first = await w.adapter.startSession(w.context);
    w.select(1);
    const second = await w.adapter.startSession(w.context);
    expect(w.factory).toHaveBeenCalledTimes(2);
    expect(w.workers[1]!.runtime.env).toMatchObject({ COPILOT_BIN: w.installs[1]!.command });
    await collect(w.adapter.sendTurn(w.input(first, "a")));
    await collect(w.adapter.sendTurn(w.input(second, "b")));
    await w.adapter.cancelTurn(first);
    expect(w.workers[0]!.turns).toHaveBeenCalledOnce();
    expect(w.workers[1]!.turns).toHaveBeenCalledOnce();
    expect(w.workers[0]!.cancelled).toHaveBeenCalledWith(first);
    expect(w.workers[1]!.cancelled).not.toHaveBeenCalled();
    expect(w.adapter.runtimeRevision()).toBe(w.installs[1]!.id);
  });

  it("does not use a CLI file that is missing or was changed, and says so instead of failing", async () => {
    const w = fixture();
    fs.writeFileSync(w.installs[0]!.command, "changed");
    expect(await w.adapter.detect()).toMatchObject({ provider: "copilot", status: "error", hint: expect.stringContaining("office-partner repair") });
    await expect(w.adapter.startSession(w.context)).rejects.toThrow("RUNTIME_MISSING");
  });

  it("reports a disabled provider as not installed", async () => {
    const w = fixture();
    store.changeState(w.dataDir, (state: any) => { state.disabled = ["copilot"]; });
    expect(await w.adapter.detect()).toMatchObject({ status: "not_installed" });
    expect(w.adapter.runtimeRevision()).toBe("disabled");
  });

  it("refuses a runtime that closes while it starts, without telling the owner it closed", async () => {
    const w = fixture();
    await w.adapter.startSession(w.context);
    const closed = vi.fn();
    w.workers[0]!.start.mockImplementationOnce(async ctx => { ctx.onRuntimeClosed?.(); return "late"; });
    const closeSession = vi.spyOn(w.workers[0]!, "closeSession");
    await expect(w.adapter.startSession({ ...w.context, onRuntimeClosed: closed })).rejects.toThrow("closed during initialization");
    expect(closeSession).toHaveBeenCalledWith("late");
    expect(closed).not.toHaveBeenCalled();
  });

  it("tells the owner when a running session's runtime closes later", async () => {
    const w = fixture();
    const closed = vi.fn();
    let notify: (() => void) | undefined;
    await w.adapter.startSession({ ...w.context, onRuntimeClosed: closed });
    w.workers[0]!.start.mockImplementationOnce(async ctx => { notify = ctx.onRuntimeClosed; return "ready"; });
    await w.adapter.startSession({ ...w.context, onRuntimeClosed: closed });
    notify?.();
    expect(closed).toHaveBeenCalledOnce();
  });

  it("uninstalls through the provider: the CLI stops being used and the provider is reported as not installed", async () => {
    const w = fixture();
    await w.adapter.uninstall();
    expect(store.readState(w.dataDir)).toMatchObject({ active: {}, disabled: ["copilot"] });
    expect(await w.adapter.detect()).toMatchObject({ status: "not_installed" });
  });

  it("has no worker for a session it did not open", () => {
    const w = fixture();
    expect(() => w.adapter.sendTurn(w.input("unknown", "x"))).toThrow("Provider worker is not available");
  });
});
