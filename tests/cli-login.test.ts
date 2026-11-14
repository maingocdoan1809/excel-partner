import { EventEmitter } from "node:events";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ClaudeAdapter } from "../providers/src/claude/adapter.js";
import { CodexAdapter } from "../providers/src/codex/adapter.js";
import { CopilotAdapter } from "../providers/src/copilot/adapter.js";
import type { ProviderAdapter, ProviderLoginAttempt, ProviderRuntime } from "../providers/src/types.js";

const { spawnMock } = vi.hoisted(() => ({ spawnMock: vi.fn() }));
vi.mock("cross-spawn", () => ({ default: spawnMock }));

function fakeChild() {
  return Object.assign(new EventEmitter(), {
    stdout: new EventEmitter(), stderr: new EventEmitter(), stdin: Object.assign(new EventEmitter(), { write: vi.fn() }),
    exitCode: null as number | null, signalCode: null, kill: vi.fn(),
  });
}

/** Starts a sign-in the way the bridge does: through the adapter, whose CLI is replaced by `login-test`. */
function startLogin(adapter: new (runtime: ProviderRuntime) => ProviderAdapter, binVar: string): ProviderLoginAttempt {
  const runtime: ProviderRuntime = { env: { [binVar]: "login-test" }, bridgeUrl: "", mcp: { command: "", args: [] }, logger: { debug() {}, warn() {} } };
  return new adapter(runtime).startLogin();
}

describe("signing in by running the provider's own CLI", () => {
  let child: ReturnType<typeof fakeChild>;
  beforeEach(() => { vi.useFakeTimers(); child = fakeChild(); spawnMock.mockReset().mockReturnValue(child); });
  afterEach(() => vi.useRealTimers());

  it("Copilot: reads split ANSI output without returning raw CLI output or credentials", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    expect(attempt.status().state).toBe("starting");
    expect(spawnMock).toHaveBeenCalledWith("login-test", ["login", "--device-code"], expect.objectContaining({ windowsHide: true }));
    child.stderr.emit("data", Buffer.from("\x1b[32mTo authenticate, visit https://github.com/login/"));
    child.stderr.emit("data", Buffer.from("device and enter code AB12-CD34\x1b[0m\nprivate-token-value"));
    expect(attempt.status()).toMatchObject({ state: "waiting", provider: "copilot", userCode: "AB12-CD34", verificationUri: "https://github.com/login/device" });
    expect(JSON.stringify(attempt.status())).not.toContain("private-token-value");
    child.exitCode = 0;
    child.emit("close", 0);
    expect(attempt.status()).toMatchObject({ state: "success", userCode: null, verificationUri: null, expiresAt: null });
  });

  it("cancels the process and ignores late output and exit events", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    attempt.cancel();
    expect(attempt.status().state).toBe("cancelled");
    expect(child.kill).toHaveBeenCalledOnce();
    child.stdout.emit("data", Buffer.from("https://github.com/login/device enter code AB12-CD34"));
    child.emit("close", 0);
    expect(attempt.status()).toMatchObject({ state: "cancelled", userCode: null });
    expect(vi.getTimerCount()).toBe(0);
  });

  it("gives up on a CLI that never produces a code", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    vi.advanceTimersByTime(30_000);
    expect(attempt.status().state).toBe("error");
    expect(child.kill).toHaveBeenCalledOnce();
  });

  it("expires a sign-in that nobody finishes", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    child.stdout.emit("data", Buffer.from("Enter one-time code: AB12-CD34 at https://github.com/login/device"));
    vi.advanceTimersByTime(30_000);
    expect(attempt.status().state).toBe("waiting");
    vi.advanceTimersByTime(10 * 60_000);
    expect(attempt.status()).toMatchObject({ state: "expired", userCode: null });
  });

  it("reports a missing CLI without exposing its output", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    child.emit("error", new Error("ENOENT private details"));
    expect(attempt.status().state).toBe("error");
    expect(attempt.status().message).not.toContain("private details");
  });

  it("does not treat an exit without a device code as a successful sign-in", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    child.exitCode = 0;
    child.emit("close", 0);
    expect(attempt.status().state).toBe("error");
  });

  it("Codex: reads the page and one-time code", () => {
    const attempt = startLogin(CodexAdapter, "CODEX_BIN");
    expect(spawnMock).toHaveBeenCalledWith("login-test", ["login", "--device-auth"], expect.objectContaining({ stdio: ["ignore", "pipe", "pipe"] }));
    child.stdout.emit("data", Buffer.from("1. Open this link\n   https://auth.openai.com/codex/device\n\n2. Enter this one-time code (expires in 15 minutes)\n   DVU8-YTCH4\n"));
    expect(attempt.status()).toMatchObject({ state: "waiting", provider: "codex", userCode: "DVU8-YTCH4", verificationUri: "https://auth.openai.com/codex/device" });
    child.exitCode = 0;
    child.emit("close", 0);
    expect(attempt.status().state).toBe("success");
  });

  it("Claude: waits for the whole link, then passes the pasted code to the CLI", () => {
    const attempt = startLogin(ClaudeAdapter, "CLAUDE_BIN");
    expect(spawnMock).toHaveBeenCalledWith("login-test", ["auth", "login"], expect.objectContaining({ stdio: ["pipe", "pipe", "pipe"] }));
    child.stdout.emit("data", Buffer.from("Opening browser\nIf the browser didn't open, visit: https://claude.com/cai/oauth/authorize?code=true&sta"));
    expect(attempt.status().state).toBe("starting");
    child.stdout.emit("data", Buffer.from("te=abc\nPaste code here if prompted > "));
    expect(attempt.status()).toMatchObject({ state: "waiting", needsCode: true, userCode: null, verificationUri: "https://claude.com/cai/oauth/authorize?code=true&state=abc" });
    expect(() => attempt.submitCode("bad code!")).toThrow("sign-in code");
    attempt.submitCode(" abc123#xyz ");
    expect(attempt.status()).toMatchObject({ state: "waiting", needsCode: false });
    expect(child.stdin.write).toHaveBeenCalledWith("abc123#xyz\n");
    expect(() => attempt.submitCode("abc123#xyz")).toThrow("no sign-in waiting");
    child.emit("close", 1);
    expect(attempt.status()).toMatchObject({ state: "error", message: expect.stringContaining("not accepted") });
  });

  it("refuses a code from a provider whose page does not hand one back", () => {
    const attempt = startLogin(CopilotAdapter, "COPILOT_BIN");
    child.stdout.emit("data", Buffer.from("https://github.com/login/device enter code AB12-CD34"));
    expect(() => attempt.submitCode("abcd1234")).toThrow("no sign-in waiting");
  });
});
