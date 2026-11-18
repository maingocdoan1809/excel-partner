import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
// @ts-expect-error The npm executable ships as JavaScript, independently of the TS build.
import { readUpdateStatus, stopBridgeForUpdate, updatePackage, writeUpdateStatus } from "../bin/update.mjs";

const dirs: string[] = [];
afterEach(() => { vi.restoreAllMocks(); for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const options = () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-update-"));
  dirs.push(dataDir);
  vi.spyOn(console, "log").mockImplementation(() => undefined);
  vi.spyOn(console, "error").mockImplementation(() => undefined);
  return { dataDir };
};

describe("update runner", () => {
  it("refuses to run without knowing the package and data folders, and leaves the checkout alone", () => {
    const before = fs.readdirSync("bin").sort();
    const result = spawnSync(process.execPath, ["bin/update-runner.mjs"], { env: { ...process.env, EP_ROOT: "", EP_DATA: "" }, encoding: "utf8" });
    expect(result.status).toBe(2);
    expect(result.stderr).toContain("EP_ROOT and EP_DATA are required");
    expect(fs.readdirSync("bin").sort()).toEqual(before);
  });
});

describe("update status file", () => {
  it("merges updates, keeps the earlier fields and never leaves a half-written file", () => {
    const { dataDir } = options();
    expect(readUpdateStatus(dataDir)).toBeNull();
    writeUpdateStatus(dataDir, { state: "running", from: "0.1.19" });
    writeUpdateStatus(dataDir, { step: "installing" });
    expect(readUpdateStatus(dataDir)).toMatchObject({ state: "running", from: "0.1.19", step: "installing", updatedAt: expect.any(String) });
    expect(fs.readdirSync(dataDir).filter((name) => name.endsWith(".tmp"))).toEqual([]);
    fs.writeFileSync(path.join(dataDir, "update-status.json"), "{ broken");
    expect(readUpdateStatus(dataDir)).toBeNull();
    expect(writeUpdateStatus(dataDir, { state: "done" })).toMatchObject({ state: "done" });
  });
});

describe("stopping the bridge for an update", () => {
  it("falls back to stopping the package processes when the gentle stop fails, for example on a token mismatch", async () => {
    const sweep = vi.fn(), warn = vi.fn();
    await stopBridgeForUpdate({ gentle: async () => { throw new Error("SERVICE_CONTROL_UNAVAILABLE (HTTP 401)"); }, sweep, warn });
    expect(sweep).toHaveBeenCalledOnce();
    expect(warn).toHaveBeenCalledWith(expect.stringContaining("HTTP 401"));
  });
  it("still refuses while the assistant is answering someone", async () => {
    const sweep = vi.fn();
    await expect(stopBridgeForUpdate({ gentle: async () => { throw new Error("The assistant is still working. Stop the response before stopping the bridge."); }, sweep })).rejects.toThrow("still working");
    expect(sweep).not.toHaveBeenCalled();
  });
});

describe("installing when a native module cannot be prepared", () => {
  it("tries once more without install scripts, and does not when the folder is only locked", async () => {
    const { dataDir } = options();
    const root = path.join(dataDir, "node_modules", "office-partner");
    fs.mkdirSync(root, { recursive: true }); fs.writeFileSync(path.join(root, "package.json"), JSON.stringify({ name: "office-partner", version: "1.0.0" }));
    const make = (firstError: string) => {
      const installs: string[][] = [];
      const run = vi.fn((_cmd: string, args: string[]) => {
        if (args[0] === "root") return { status: 0, stdout: path.join(dataDir, "node_modules") };
        if (args[0] === "install") { installs.push(args); return installs.length === 1 ? { status: 1, stderr: firstError } : { status: 0, stderr: "" }; }
        return { status: 0 };
      });
      return { run, installs };
    };
    const script = make("npm error gyp ERR! build error better-sqlite3");
    await updatePackage({ root, dataDir, nodePath: "node", stop: async () => {}, wasRunning: false, run: script.run as never, spec: "office-partner@1" });
    expect(script.installs).toHaveLength(2); expect(script.installs[1]).toContain("--ignore-scripts"); expect(script.installs[0]).not.toContain("--ignore-scripts");
    const locked = make("npm error code EBUSY");
    await updatePackage({ root, dataDir, nodePath: "node", stop: async () => {}, wasRunning: false, run: locked.run as never, spec: "office-partner@1", wait: async () => {} });
    expect(locked.installs.every(args => !args.includes("--ignore-scripts"))).toBe(true);
  });
});
