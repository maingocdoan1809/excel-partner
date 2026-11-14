import { spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { reapRecordedProcesses, setProcessRegistryFile, trackProcess } from "../providers/src/process-registry.js";
import { crashSummary } from "../shared/src/crash-exit.js";

let dir = "";
afterEach(() => { setProcessRegistryFile(null); if (dir) fs.rmSync(dir, { recursive: true, force: true }); });
const temp = () => (dir = fs.mkdtempSync(path.join(os.tmpdir(), "process-registry-")));

describe("provider process registry", () => {
  it("records a provider CLI while it runs and forgets it when it exits", async () => {
    const file = path.join(temp(), "provider-processes.json");
    setProcessRegistryFile(file);
    const child = spawn(process.execPath, ["-e", "setTimeout(() => {}, 200)"], { stdio: "ignore" });
    trackProcess(child);
    const [record] = JSON.parse(fs.readFileSync(file, "utf8"));
    expect(record).toMatchObject({ pid: child.pid, exe: path.basename(process.execPath) });
    expect(Math.abs(record.startedAt - Date.now())).toBeLessThan(5000);
    await new Promise(resolve => child.once("exit", resolve));
    expect(fs.existsSync(file)).toBe(false);
  });

  it.runIf(process.platform === "win32")("stops a recorded process only when its executable and start time still match", () => {
    const file = path.join(temp(), "provider-processes.json");
    const startedAt = 1_700_000_000_000;
    fs.writeFileSync(file, JSON.stringify([
      { pid: 101, startedAt, exe: "cmd.exe" },        // still ours
      { pid: 102, startedAt, exe: "codex.exe" },      // number reused by another program
      { pid: 103, startedAt, exe: "node.exe" },       // same program name, started later: a reused number
      { pid: 104, startedAt, exe: "opencode.exe" },   // gone
    ]));
    const observed = [{ pid: 101, name: "cmd.exe", start: startedAt + 40 }, { pid: 102, name: "OneDrive.exe", start: startedAt }, { pid: 103, name: "node.exe", start: startedAt + 600_000 }];
    const run = vi.fn((command: string) => command === "powershell.exe" ? { status: 0, stdout: JSON.stringify(observed) } : { status: 0, stdout: "" });
    expect(reapRecordedProcesses(file, run as never)).toBe(1);
    expect(run).toHaveBeenCalledWith("taskkill", ["/PID", "101", "/T", "/F"], expect.anything());
    expect(run.mock.calls.filter(([command]) => command === "taskkill")).toHaveLength(1);
    expect(fs.existsSync(file)).toBe(false);
  });

  it("does nothing without a record, and kills nothing when the lookup fails", () => {
    const file = path.join(temp(), "provider-processes.json");
    const run = vi.fn((_command: string) => ({ status: 1, stdout: "" }));
    expect(reapRecordedProcesses(file, run as never)).toBe(0);
    fs.writeFileSync(file, JSON.stringify([{ pid: 101, startedAt: 1, exe: "cmd.exe" }]));
    expect(reapRecordedProcesses(file, run as never)).toBe(0);
    expect(run.mock.calls.some(([command]) => command === "taskkill")).toBe(false);
  });
});

describe("crash log", () => {
  it("keeps the error type, code and stack frames but not the message, which can quote user content", () => {
    const error = Object.assign(new Error("prompt: secret quarterly numbers"), { code: "EPIPE" });
    const summary = crashSummary(error);
    expect(summary).toMatchObject({ name: "Error", code: "EPIPE" });
    expect(summary.frames?.length).toBeGreaterThan(0);
    expect(JSON.stringify(summary)).not.toContain("secret");
    expect(crashSummary("text with secret")).toEqual({ name: "string" });
  });
});
