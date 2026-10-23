import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import type { ProviderEvent } from "../protocol/src/index.js";

const cli = vi.hoisted(() => ({ spawn: vi.fn(), messages: [] as object[], code: 0 }));
vi.mock("cross-spawn", () => ({ default: cli.spawn }));
import { ClaudeAdapter } from "../providers/src/claude.js";

const runtime = { bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } };
let adapter: ClaudeAdapter;
beforeEach(() => {
  adapter = new ClaudeAdapter(runtime);
  vi.spyOn(adapter, "sessionTitle").mockResolvedValue(undefined);
  cli.messages = [];
  cli.code = 0;
  cli.spawn.mockReset().mockImplementation(() => {
    const child = Object.assign(new EventEmitter(), {
      stdin: new PassThrough(), stdout: new PassThrough(), stderr: new PassThrough(),
      exitCode: null as number | null, signalCode: null,
      kill: vi.fn(),
    });
    setImmediate(() => {
      child.stdout.end(cli.messages.map(message => JSON.stringify(message)).join("\n") + "\n");
      child.stderr.end();
      child.exitCode = cli.code;
      child.emit("exit", cli.code, null);
    });
    return child;
  });
});
afterEach(async () => { await adapter.dispose(); vi.restoreAllMocks(); });

async function start(resumeSessionId?: string, webSearch = false) {
  return adapter.startSession({ workbookConnectionId: "workbook", sessionId: "bridge", resumeSessionId,
    signal: new AbortController().signal,
    mcp: { command: process.execPath, args: [], env: webSearch ? { EXCEL_WEB_SEARCH: "1" } : {}, bridgeUrl: "", sessionId: "", token: "" },
  });
}
async function turn(providerSessionId: string, signal = new AbortController().signal, systemPrompt?: string) {
  const events: ProviderEvent[] = [];
  for await (const event of adapter.sendTurn({ providerSessionId, requestId: "request", prompt: "hello", selection: null, signal, systemPrompt })) {
    events.push(event);
  }
  return events;
}

describe("Claude native session confirmation", () => {
  it.each([false, true])("exposes native web tools only when Web search is enabled (%j)", async enabled => {
    await turn(await start(undefined, enabled));
    const args = cli.spawn.mock.calls[0]![1] as string[];
    expect(args[args.indexOf("--tools") + 1]).toBe(enabled ? "WebSearch,WebFetch" : "");
    expect(args.includes("WebSearch")).toBe(enabled);
    expect(args.includes("WebFetch")).toBe(enabled);
  });

  it.each([undefined, "", " ", "Custom instructions"])("only overrides native instructions for a non-empty add-in prompt (%j)", async systemPrompt => {
    await turn(await start(), new AbortController().signal, systemPrompt);
    const args = cli.spawn.mock.calls[0]![1] as string[];
    if (systemPrompt?.trim()) {
      expect(args[args.indexOf("--system-prompt") + 1]).toBe(systemPrompt);
    } else {
      expect(args).not.toContain("--system-prompt");
    }
  });

  it("does not publish a locally generated ID when cancelled before launch", async () => {
    const id = await start();
    const controller = new AbortController();
    controller.abort();
    const events = await turn(id, controller.signal);
    expect(events.some(event => event.type === "session_info")).toBe(false);
    expect(events.at(-1)).toMatchObject({ reason: "cancelled" });
    expect(cli.spawn).not.toHaveBeenCalled();
  });

  it.each([[], [{ type: "result", is_error: true, result: "startup failed" }]])(
    "keeps an unconfirmed session fresh after CLI startup failure (%j)", async (...messages) => {
      cli.messages = messages.flat();
      cli.code = 1;
      const id = await start();
      expect((await turn(id)).some(event => event.type === "session_info")).toBe(false);
      await turn(id);
      for (const [, args] of cli.spawn.mock.calls) {
        expect(args).toContain("--session-id");
        expect(args).not.toContain("--resume");
      }
    },
  );

  it.each(["init", "result"])("publishes and resumes the CLI-confirmed ID from %s", async kind => {
    cli.messages = [kind === "init"
      ? { type: "system", subtype: "init", session_id: "confirmed-session" }
      : { type: "result", session_id: "confirmed-session", is_error: false }];
    const id = await start();
    expect((await turn(id)).find(event => event.type === "session_info"))
      .toMatchObject({ providerSessionId: "confirmed-session" });
    cli.messages = [{ type: "result", session_id: "confirmed-session", is_error: false }];
    await turn(id);
    const args = cli.spawn.mock.calls[1]![1] as string[];
    expect(args[args.indexOf("--resume") + 1]).toBe("confirmed-session");
    expect(args).not.toContain("--session-id");
  });

  it("publishes an existing resume ID without waiting for a new CLI confirmation", async () => {
    const events = await turn(await start("existing-session"));
    expect(events[0]).toMatchObject({ type: "session_info", providerSessionId: "existing-session" });
  });
});
