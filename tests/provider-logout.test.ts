import { EventEmitter } from "node:events";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ClaudeAdapter } from "../providers/src/claude/adapter.js";
import { CodexAdapter } from "../providers/src/codex/adapter.js";
import { CopilotAdapter } from "../providers/src/copilot/adapter.js";
import { OpencodeAdapter } from "../providers/src/opencode/adapter.js";
import { ProviderLoginError } from "../providers/src/login.js";
import type { ProviderAdapter, ProviderRuntime } from "../providers/src/types.js";

const { spawnMock } = vi.hoisted(() => ({ spawnMock: vi.fn() }));
vi.mock("cross-spawn", () => ({ default: spawnMock }));

/** A CLI that finishes at once with the given exit code. */
function finishing(code: number) {
  const child = Object.assign(new EventEmitter(), { stdout: new EventEmitter(), stderr: new EventEmitter(), stdin: null, exitCode: null, signalCode: null, kill: vi.fn() });
  queueMicrotask(() => child.emit("close", code));
  return child;
}
const adapter = (kind: new (runtime: ProviderRuntime) => ProviderAdapter, binVar: string) =>
  new kind({ env: { [binVar]: "logout-test" }, bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } });

describe("signing a provider out", () => {
  beforeEach(() => spawnMock.mockReset().mockImplementation(() => finishing(0)));

  it.each([
    ["Claude", ClaudeAdapter, "CLAUDE_BIN", ["auth", "logout"]],
    ["Codex", CodexAdapter, "CODEX_BIN", ["logout"]],
  ] as const)("%s runs its own sign-out command", async (_name, kind, binVar, args) => {
    await adapter(kind, binVar).logout();
    expect(spawnMock).toHaveBeenCalledWith("logout-test", [...args], expect.anything());
  });

  it("reports a sign-out that did not work, without the CLI's own output", async () => {
    spawnMock.mockImplementation(() => finishing(1));
    await expect(adapter(CodexAdapter, "CODEX_BIN").logout()).rejects.toBeInstanceOf(ProviderLoginError);
  });

  it("OpenCode signs out the account it is told, and only a plain account id", async () => {
    const opencode = adapter(OpencodeAdapter, "OPENCODE_BIN");
    await expect(opencode.logout()).rejects.toThrow("Choose the account");
    await expect(opencode.logout("../x")).rejects.toThrow("Choose the account");
    await opencode.logout("openai");
    expect(spawnMock).toHaveBeenCalledWith("logout-test", ["providers", "logout", "openai"], expect.anything());
  });

  it("OpenCode stops its sign-in service around a sign-out, so the next account list is read fresh", async () => {
    const opencode = adapter(OpencodeAdapter, "OPENCODE_BIN");
    const service = { closed: false, close: vi.fn(() => { service.closed = true; }) };
    (opencode as unknown as { auth: typeof service }).auth = service;
    await opencode.logout("openai");
    expect(service.close).toHaveBeenCalled();
    expect(service.closed).toBe(true);
  });

  it("Copilot has no sign-out command, and says so in its capabilities", async () => {
    expect(CopilotAdapter.capabilities.logout).toBe(false);
    await expect(adapter(CopilotAdapter, "COPILOT_BIN").logout()).rejects.toBeInstanceOf(ProviderLoginError);
    expect(spawnMock).not.toHaveBeenCalled();
    expect([ClaudeAdapter, CodexAdapter, OpencodeAdapter].every(kind => kind.capabilities.logout)).toBe(true);
  });
});
