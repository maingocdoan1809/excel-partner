import { EventEmitter } from "node:events";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ProviderLoginManager } from "../bridge/src/provider-login.js";

const { spawnMock } = vi.hoisted(() => ({ spawnMock: vi.fn() }));
vi.mock("cross-spawn", () => ({ default: spawnMock }));

function fakeChild() {
  return Object.assign(new EventEmitter(), {
    stdout: new EventEmitter(), stderr: new EventEmitter(),
    exitCode: null as number | null, signalCode: null, kill: vi.fn(),
  });
}

describe("Copilot device-code login", () => {
  let manager: ProviderLoginManager;
  let child: ReturnType<typeof fakeChild>;

  beforeEach(() => {
    vi.useFakeTimers();
    child = fakeChild();
    spawnMock.mockReset().mockReturnValue(child);
    manager = new ProviderLoginManager("copilot-test");
  });

  afterEach(() => { manager.dispose(); vi.useRealTimers(); });

  it("parses split ANSI output without returning raw CLI output or credentials", () => {
    expect(manager.start("workbook").state).toBe("starting");
    expect(spawnMock).toHaveBeenCalledWith("copilot-test", ["login", "--device-code"], expect.objectContaining({ windowsHide: true }));
    child.stderr.emit("data", Buffer.from("\x1b[32mTo authenticate, visit https://github.com/login/"));
    child.stderr.emit("data", Buffer.from("device and enter code AB12-CD34\x1b[0m\nprivate-token-value"));
    expect(manager.status("workbook")).toMatchObject({ state: "waiting", userCode: "AB12-CD34", verificationUri: "https://github.com/login/device" });
    expect(JSON.stringify(manager.status("workbook"))).not.toContain("private-token-value");
    child.exitCode = 0;
    child.emit("close", 0);
    expect(manager.status("workbook")).toMatchObject({ state: "success", userCode: null, verificationUri: null, expiresAt: null });
  });

  it("isolates workbook sessions and prevents competing login processes", () => {
    manager.start("first");
    manager.start("first");
    expect(spawnMock).toHaveBeenCalledTimes(1);
    expect(() => manager.start("second")).toThrow("Another Excel window");
    expect(manager.status("second").state).toBe("idle");
    manager.cancel("second");
    expect(manager.status("first").state).toBe("starting");
  });

  it("cancels a process and ignores late output and exit events", () => {
    manager.start("workbook");
    expect(manager.cancel("workbook").state).toBe("cancelled");
    expect(child.kill).toHaveBeenCalledOnce();
    child.stdout.emit("data", Buffer.from("https://github.com/login/device enter code AB12-CD34"));
    child.emit("close", 0);
    expect(manager.status("workbook")).toMatchObject({ state: "cancelled", userCode: null });
  });

  it("times out a CLI that never produces a code", () => {
    manager.start("workbook");
    vi.advanceTimersByTime(30_000);
    expect(manager.status("workbook").state).toBe("error");
    expect(child.kill).toHaveBeenCalledOnce();
  });

  it("expires waiting sessions and permits a fresh attempt", () => {
    manager.start("workbook");
    child.stdout.emit("data", Buffer.from("Enter one-time code: AB12-CD34 at https://github.com/login/device"));
    vi.advanceTimersByTime(30_000);
    expect(manager.status("workbook").state).toBe("waiting");
    vi.advanceTimersByTime(10 * 60_000);
    expect(manager.status("workbook")).toMatchObject({ state: "expired", userCode: null });
    expect(manager.start("workbook").state).toBe("starting");
  });

  it("handles missing or unsupported CLIs without exposing their output", () => {
    manager.start("workbook");
    child.emit("error", new Error("ENOENT private details"));
    expect(manager.status("workbook").state).toBe("error");
    expect(manager.status("workbook").message).not.toContain("private details");
  });

  it("does not treat an exit without a device code as successful authentication", () => {
    manager.start("workbook");
    child.exitCode = 0;
    child.emit("close", 0);
    expect(manager.status("workbook").state).toBe("error");
  });

  it("forgets login codes and kills pending processes on disconnect", () => {
    manager.start("workbook");
    child.stdout.emit("data", Buffer.from("https://github.com/login/device enter code AB12-CD34"));
    manager.forget("workbook");
    expect(manager.status("workbook").state).toBe("idle");
    expect(child.kill).toHaveBeenCalledOnce();
    expect(vi.getTimerCount()).toBe(0);
  });
});
