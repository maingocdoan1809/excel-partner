import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
// @ts-expect-error Ships as plain JavaScript beside the other bin tools.
import { recordStartup, startupCode } from "../bin/startup-log.mjs";
// @ts-expect-error Ships as plain JavaScript beside the other bin tools.
import { readLastStartup as lastStartup } from "../bin/doctor-probes.mjs";
// @ts-expect-error Ships as plain JavaScript beside the other bin tools.
import { managementMain } from "../bin/manage.mjs";

const dirs: string[] = [];
// The launcher repeats the start command's error text on its own stderr; keep the test output readable.
beforeEach(() => { vi.spyOn(process.stderr, "write").mockImplementation(() => true); });
const folder = () => { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "startup-log-")); dirs.push(dir); return dir; };
afterEach(() => { process.exitCode = undefined; vi.restoreAllMocks(); for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const entries = (dir: string) => fs.readFileSync(path.join(dir, "startup.log"), "utf8").trim().split("\n").map(line => JSON.parse(line));

describe("startup log", () => {
  it("keeps one line per attempt and finds the newest for each service", () => {
    const dir = folder();
    recordStartup(dir, { service: "addin", attempt: 1, ok: false, code: "NOT_ANSWERING", message: "first line\nsecond line" });
    recordStartup(dir, { service: "gateway", attempt: 1, ok: true });
    recordStartup(dir, { service: "addin", attempt: 2, ok: true });
    expect(entries(dir)).toHaveLength(3);
    expect(entries(dir)[0]).toMatchObject({ service: "addin", ok: false, message: "first line" });
    expect(lastStartup(dir, "addin")).toMatchObject({ attempt: 2, ok: true });
    expect(lastStartup(dir, "gateway")).toMatchObject({ ok: true });
    expect(lastStartup(folder(), "addin")).toBeNull();
  });

  it("stays small and survives a damaged line or an unwritable folder", () => {
    const dir = folder();
    for (let i = 0; i < 400; i++) recordStartup(dir, { service: "addin", attempt: 1, ok: true, message: "x".repeat(200) });
    expect(fs.statSync(path.join(dir, "startup.log")).size).toBeLessThan(40 * 1024);
    fs.appendFileSync(path.join(dir, "startup.log"), "{broken\n");
    expect(lastStartup(dir, "addin")).toMatchObject({ ok: true });
    const blocked = path.join(folder(), "file"); fs.writeFileSync(blocked, "");
    expect(() => recordStartup(path.join(blocked, "inside"), { service: "addin", ok: true })).not.toThrow();
  });

  it("names the failure without copying text that may hold a user name or path", () => {
    expect(startupCode("Bridge process 7 is running but not answering.")).toBe("NOT_ANSWERING");
    expect(startupCode("Recorded gateway is still alive but unhealthy")).toBe("NOT_ANSWERING");
    expect(startupCode("Error: listen EADDRINUSE: address already in use 127.0.0.1:8765")).toBe("PORT_IN_USE");
    expect(startupCode("Bridge did not start. See C:\\Users\\someone\\bridge.err.log")).toBe("START_FAILED");
    expect(startupCode("Run 'office-partner setup' first.")).toBe("NOT_SET_UP");
    expect(startupCode("ACTIVE_RELEASE_DAMAGED_USE_RECOVERY")).toBe("RELEASE_DAMAGED");
    expect(startupCode(undefined)).toBe("START_FAILED");
  });
});

describe("start from the sign-in entry and from Fix connection", () => {
  const options = (dir: string, launch: (root: string, args: string[]) => { status: number; stderr: string }) => ({ dataDir: dir, root: dir, launch, startPauseMs: 1 });

  it("D3 tries a failed start again and records every attempt, so a failure is never silent", async () => {
    const dir = folder();
    const launch = vi.fn<(root: string, args: string[]) => { status: number; stderr: string }>()
      .mockReturnValueOnce({ status: 1, stderr: "Bridge process 7 is running but not answering." })
      .mockReturnValueOnce({ status: 1, stderr: "Error: listen EADDRINUSE" })
      .mockReturnValueOnce({ status: 0, stderr: "" });
    await managementMain(["start", "--at-logon"], options(dir, launch));
    expect(launch).toHaveBeenCalledTimes(3);
    expect(launch.mock.calls[0]![1]).toEqual(["start"]); // the sign-in flag is not passed on
    expect(entries(dir).map(entry => [entry.service, entry.attempt, entry.ok, entry.code])).toEqual([
      ["addin", 1, false, "NOT_ANSWERING"], ["addin", 2, false, "PORT_IN_USE"], ["addin", 3, true, undefined]]);
    expect(process.exitCode).toBeUndefined();
  });

  it("gives up after three attempts and reports failure", async () => {
    const dir = folder();
    const launch = vi.fn((_root: string, _args: string[]) => ({ status: 1, stderr: "Bridge did not start." }));
    await managementMain(["start", "gateway", "--at-logon"], options(dir, launch));
    expect(launch).toHaveBeenCalledTimes(3);
    expect(launch.mock.calls[0]![1]).toEqual(["start", "gateway"]);
    expect(entries(dir).every(entry => entry.service === "gateway" && entry.ok === false)).toBe(true);
    expect(process.exitCode).toBe(1);
  });

  it("does not repeat an interactive start", async () => {
    const dir = folder();
    const launch = vi.fn((_root: string, _args: string[]) => ({ status: 1, stderr: "Bridge did not start." }));
    await managementMain(["start"], options(dir, launch));
    expect(launch).toHaveBeenCalledTimes(1);
  });

  it("D5 Fix connection tries to reinstall a damaged release once, then reports what is still wrong", async () => {
    const dir = folder(); const gatewayHome = folder();
    vi.stubEnv("AGENT_GATEWAY_HOME", gatewayHome);
    const log = vi.spyOn(console, "log").mockImplementation(() => {});
    try {
      const launch = vi.fn((_root: string, _args: string[]) => ({ status: 1, stderr: "ACTIVE_RELEASE_DAMAGED_USE_RECOVERY" }));
      await managementMain(["fix"], options(dir, launch));
      const said = log.mock.calls.flat().join("\n");
      expect(said).toContain("installed files are damaged");
      // Nothing is installed in this empty folder, so the repair cannot run; the person is told so and the log keeps the reason.
      expect(said).toContain("FAILED Office Partner");
      expect(said).toContain("startup.log");
      expect(launch).toHaveBeenCalledTimes(1);
      expect(entries(dir).at(-1)).toMatchObject({ service: "addin", ok: false, source: "fix" });
      expect(process.exitCode).toBe(1);
    } finally { vi.unstubAllEnvs(); }
  });

  it("D4 Fix connection restarts the bridge with repair, adds the gateway only when it is set up, and says what happened", async () => {
    const dir = folder(); const gatewayHome = folder();
    vi.stubEnv("AGENT_GATEWAY_HOME", gatewayHome);
    const log = vi.spyOn(console, "log").mockImplementation(() => {});
    try {
      const launch = vi.fn((_root: string, _args: string[]) => ({ status: 0, stderr: "" }));
      await managementMain(["fix"], options(dir, launch));
      expect(launch.mock.calls.map(call => call[1])).toEqual([["start", "--repair"]]);
      fs.writeFileSync(path.join(gatewayHome, "token.txt"), "token");
      const failing = vi.fn((_root: string, args: string[]) => args.includes("gateway") ? { status: 1, stderr: "Recorded gateway is still alive but unhealthy" } : { status: 0, stderr: "" });
      await managementMain(["fix"], options(dir, failing));
      expect(failing.mock.calls.map(call => call[1])).toEqual([["start", "--repair"], ["start", "gateway", "--repair"]]);
      const said = log.mock.calls.flat().join("\n");
      expect(said).toContain("OK     Office Partner");
      expect(said).toContain("FAILED Gateway");
      expect(said).toContain("NOT_ANSWERING");
      expect(said).not.toContain(gatewayHome);
      expect(entries(dir).some(entry => entry.source === "fix" && entry.service === "gateway" && entry.ok === false)).toBe(true);
      expect(process.exitCode).toBe(1);
    } finally { vi.unstubAllEnvs(); }
  });
});
