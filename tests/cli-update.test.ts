import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
// @ts-expect-error The npm executable ships as JavaScript, independently of the TS build.
import { installedVersion, readUpdateStatus, stopPackageProcesses, updatePackage, writeUpdateStatus } from "../bin/update.mjs";

const dirs: string[] = [];
afterEach(() => { vi.restoreAllMocks(); for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const options = () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-update-"));
  dirs.push(dataDir);
  vi.spyOn(console, "log").mockImplementation(() => undefined);
  vi.spyOn(console, "error").mockImplementation(() => undefined);
  const globalRoot = path.join(dataDir, "global modules");
  const root = path.join(globalRoot, "office-partner");
  return { dataDir, root, globalRoot, nodePath: process.execPath, stop: vi.fn(async () => undefined), wasRunning: true };
};

describe("CLI update", () => {
  it("stops before installing, uses a stable working directory, then runs setup from the new package", async () => {
    const opts = options();
    const run = vi.fn()
      .mockReturnValueOnce({ status: 0, stdout: opts.globalRoot + "\n" })
      .mockImplementationOnce(() => {
        expect(opts.stop).toHaveBeenCalledOnce();
        return { status: 0 };
      })
      .mockReturnValueOnce({ status: 0 });
    await updatePackage({ ...opts, run });
    expect(run.mock.calls[1]).toEqual(["npm.cmd", ["install", "--global", "office-partner@latest"], expect.objectContaining({ cwd: opts.dataDir, shell: true })]);
    expect(run.mock.calls[2]).toEqual([process.execPath, [path.join(opts.root, "bin", "office-partner.mjs"), "setup"], expect.objectContaining({ cwd: opts.dataDir })]);
  });

  it("leaves a source or npx installation running instead of replacing another global installation", async () => {
    const opts = options();
    const run = vi.fn().mockReturnValue({ status: 0, stdout: path.join(opts.dataDir, "another-install") });
    await expect(updatePackage({ ...opts, run })).rejects.toThrow(/global npm installation/);
    expect(opts.stop).not.toHaveBeenCalled();
    expect(run).toHaveBeenCalledOnce();
  });

  it("attempts recovery when npm fails and reports setup failures distinctly", async () => {
    const opts = options();
    const failedInstall = vi.fn().mockReturnValueOnce({ status: 0, stdout: opts.globalRoot })
      .mockReturnValueOnce({ status: 1 }).mockReturnValueOnce({ status: 0 });
    await expect(updatePackage({ ...opts, run: failedInstall })).rejects.toThrow(/npm install/);
    expect(failedInstall.mock.calls[2]?.[1]).toEqual([path.join(opts.root, "bin", "office-partner.mjs"), "start"]);
    const failedSetup = vi.fn().mockReturnValueOnce({ status: 0, stdout: opts.globalRoot })
      .mockReturnValueOnce({ status: 0 }).mockReturnValueOnce({ status: 1 });
    await expect(updatePackage({ ...opts, run: failedSetup })).rejects.toThrow(/Package updated, but setup failed/);
  });

  it("installs a given spec or tarball instead of the latest release, quoting paths with spaces", async () => {
    const opts = options();
    const run = vi.fn().mockReturnValue({ status: 0, stdout: opts.globalRoot });
    await updatePackage({ ...opts, run, spec: "C:\\My Files\\excel-partner-0.2.0.tgz" });
    expect(run.mock.calls[1]![1]).toEqual(["install", "--global", '"C:\\My Files\\excel-partner-0.2.0.tgz"']);
    run.mockClear().mockReturnValue({ status: 0, stdout: opts.globalRoot });
    await updatePackage({ ...opts, run, spec: "office-partner@0.1.20" });
    expect(run.mock.calls[1]![1]).toEqual(["install", "--global", "office-partner@0.1.20"]);
  });

  it("reports each step in order, and a recovery when npm fails", async () => {
    const opts = options();
    const steps: string[] = [];
    const ok = vi.fn().mockReturnValue({ status: 0, stdout: opts.globalRoot });
    await updatePackage({ ...opts, run: ok, report: (s: string) => steps.push(s) });
    expect(steps).toEqual(["stopping", "installing", "starting"]);
    steps.length = 0;
    const failing = vi.fn().mockReturnValueOnce({ status: 0, stdout: opts.globalRoot }).mockReturnValueOnce({ status: 1 }).mockReturnValue({ status: 0 });
    await expect(updatePackage({ ...opts, run: failing, report: (s: string) => steps.push(s) })).rejects.toThrow(/npm install/);
    expect(steps).toEqual(["stopping", "installing", "recovering"]);
  });

  it("returns the version that was installed", async () => {
    const opts = options();
    fs.mkdirSync(opts.root, { recursive: true });
    fs.writeFileSync(path.join(opts.root, "package.json"), JSON.stringify({ version: "0.1.20" }));
    const run = vi.fn().mockReturnValue({ status: 0, stdout: opts.globalRoot });
    expect(await updatePackage({ ...opts, run })).toMatchObject({ version: "0.1.20", installedRoot: opts.root });
    expect(installedVersion(opts.root)).toBe("0.1.20");
    expect(installedVersion(path.join(opts.dataDir, "nothing"))).toBeNull();
  });
});

describe("stopPackageProcesses", () => {
  it("does nothing off Windows", () => {
    const run = vi.fn();
    expect(stopPackageProcesses({ root: "/opt/excel-partner", run, platform: "linux" })).toEqual({ stopped: 0 });
    expect(run).not.toHaveBeenCalled();
  });

  it("stops what runs from the package folder, sparing itself and the given pids, and reports how many", () => {
    const run = vi.fn().mockReturnValue({ status: 0, stdout: "4\r\n" });
    const result = stopPackageProcesses({ root: "C:\\Users\\o'brien\\npm\\node_modules\\excel-partner", keep: [111], run, platform: "win32" });
    expect(result).toEqual({ stopped: 4 });
    const [exe, args] = run.mock.calls[0]!;
    expect(exe).toBe("powershell.exe");
    expect(args.slice(0, 3)).toEqual(["-NoProfile", "-NonInteractive", "-EncodedCommand"]);
    const script = Buffer.from(args[3] as string, "base64").toString("utf16le");
    expect(script).toContain("'C:\\Users\\o''brien\\npm\\node_modules\\excel-partner'"); // a quote in the path cannot end the literal
    expect(script).toContain(`$spared = @(${process.pid},111)`);
    expect(script).toContain("Spare $id");              // itself and the kept pids are never stopped
    expect(script).toContain("CommandLine.IndexOf($root"); // named on the command line
    expect(script).toContain("[EpCwd]::Get");            // working directory inside the folder
    expect(script).toContain("ParentProcessId");         // descendants go too
    expect(script).toContain("Stop-Process");
  });

  it("copes with PowerShell printing nothing useful", () => {
    expect(stopPackageProcesses({ root: "C:\\x", run: vi.fn().mockReturnValue({ stdout: "" }), platform: "win32" })).toEqual({ stopped: 0 });
  });
});

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
