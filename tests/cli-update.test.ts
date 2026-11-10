import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
// @ts-expect-error The npm executable ships as JavaScript, independently of the TS build.
import { installedVersion, readUpdateStatus, removeNpmLeftovers, stopPackageProcesses, updatePackage, writeUpdateStatus } from "../bin/update.mjs";

// What every install adds: replace files npm did not write, and ask the registry for the latest version, not a cached answer.
const FORCE = ["--force", "--prefer-online", "--no-audit", "--no-fund"];

const dirs: string[] = [];
afterEach(() => { vi.restoreAllMocks(); for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const options = () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-update-"));
  dirs.push(dataDir);
  vi.spyOn(console, "log").mockImplementation(() => undefined);
  vi.spyOn(console, "error").mockImplementation(() => undefined);
  const globalRoot = path.join(dataDir, "global modules");
  const root = path.join(globalRoot, "excel-partner");
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
    expect(run.mock.calls[1]).toEqual(["npm.cmd", ["install", "--global", ...FORCE, "excel-partner@latest"], expect.objectContaining({ cwd: opts.dataDir, shell: true })]);
    expect(run.mock.calls[2]).toEqual([process.execPath, [path.join(opts.root, "bin", "excel-partner.mjs"), "setup"], expect.objectContaining({ cwd: opts.dataDir })]);
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
    expect(failedInstall.mock.calls[2]?.[1]).toEqual([path.join(opts.root, "bin", "excel-partner.mjs"), "start"]);
    const failedSetup = vi.fn().mockReturnValueOnce({ status: 0, stdout: opts.globalRoot })
      .mockReturnValueOnce({ status: 0 }).mockReturnValueOnce({ status: 1 });
    await expect(updatePackage({ ...opts, run: failedSetup })).rejects.toThrow(/Package updated, but setup failed/);
  });

  it("installs a given spec or tarball instead of the latest release, quoting paths with spaces", async () => {
    const opts = options();
    const run = vi.fn().mockReturnValue({ status: 0, stdout: opts.globalRoot });
    await updatePackage({ ...opts, run, spec: "C:\\My Files\\excel-partner-0.2.0.tgz" });
    expect(run.mock.calls[1]![1]).toEqual(["install", "--global", ...FORCE, '"C:\\My Files\\excel-partner-0.2.0.tgz"']);
    run.mockClear().mockReturnValue({ status: 0, stdout: opts.globalRoot });
    await updatePackage({ ...opts, run, spec: "excel-partner@0.1.20" });
    expect(run.mock.calls[1]![1]).toEqual(["install", "--global", ...FORCE, "excel-partner@0.1.20"]);
  });

  it("closes everything again and retries when npm reports a lock, clearing what a failed run left behind", async () => {
    const opts = options();
    fs.mkdirSync(opts.globalRoot, { recursive: true });
    const leftover = path.join(opts.globalRoot, ".excel-partner-AbC123");
    fs.mkdirSync(leftover);
    const wait = vi.fn(async () => undefined);
    const stderr = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    const run = vi.fn()
      .mockReturnValueOnce({ status: 0, stdout: opts.globalRoot })
      .mockImplementationOnce(() => { expect(fs.existsSync(leftover)).toBe(false); return { status: 1, stderr: "npm error code EBUSY\nnpm error resource busy or locked, rename" }; })
      .mockReturnValueOnce({ status: 0 })
      .mockReturnValueOnce({ status: 0 });
    await updatePackage({ ...opts, run, wait });
    expect(run.mock.calls.filter(call => call[0] === "npm.cmd" && call[1][0] === "install")).toHaveLength(2);
    expect(opts.stop).toHaveBeenCalledTimes(2);
    expect(wait).toHaveBeenCalledWith(4000);
    expect(stderr).toHaveBeenCalledWith(expect.stringContaining("EBUSY"));
    expect(run.mock.calls[3]?.[1]).toEqual([path.join(opts.root, "bin", "excel-partner.mjs"), "setup"]);
  });
  it("gives up after three locked attempts, restarts the old bridge and says what to close", async () => {
    const opts = options();
    vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    const run = vi.fn().mockReturnValueOnce({ status: 0, stdout: opts.globalRoot })
      .mockReturnValue({ status: 1, stderr: "npm error code EPERM" });
    await expect(updatePackage({ ...opts, run, wait: async () => undefined })).rejects.toThrow(/EBUSY or EPERM/);
    expect(run.mock.calls.filter(call => call[1][0] === "install")).toHaveLength(3);
    expect(opts.stop).toHaveBeenCalledTimes(3);
    expect(run.mock.calls.at(-1)?.[1]).toEqual([path.join(opts.root, "bin", "excel-partner.mjs"), "start"]);
  });
  it("does not retry what trying again cannot cure", async () => {
    const opts = options();
    vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    const run = vi.fn().mockReturnValueOnce({ status: 0, stdout: opts.globalRoot })
      .mockReturnValueOnce({ status: 1, stderr: "npm error code E404\nnpm error 404 Not Found" }).mockReturnValueOnce({ status: 0 });
    await expect(updatePackage({ ...opts, run, wait: async () => undefined })).rejects.toThrow(/npm install/);
    expect(run.mock.calls.filter(call => call[1][0] === "install")).toHaveLength(1);
    expect(opts.stop).toHaveBeenCalledOnce();
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
    expect(script).toContain(`'${path.resolve("C:\\Users\\o'brien\\npm\\node_modules\\excel-partner").replace(/'/g, "''")}'`); // a quote in the path cannot end the literal
    expect(script).toContain(`$spared = @(${process.pid},111)`);
    expect(script).toContain("Spare $id");              // itself and the kept pids are never stopped
    expect(script).toContain("CommandLine.IndexOf($root"); // named on the command line
    expect(script).toContain("[EpCwd]::Get");            // working directory inside the folder
    expect(script).toContain("ParentProcessId");         // descendants go too
    expect(script).toContain("Stop-Process");
  });

  it("also stops an update or an npm install of this package that is already running, but never the shell that ran this one", () => {
    const run = vi.fn().mockReturnValue({ status: 0, stdout: "2\r\n" });
    stopPackageProcesses({ root: "C:\\npm\\node_modules\\excel-partner", packageName: "excel-partner", run, platform: "win32" });
    const script = Buffer.from(run.mock.calls[0]![1][3] as string, "base64").toString("utf16le");
    expect(script).toContain("$pkg = 'excel-partner'");
    expect(script).toContain("update-runner");                          // another update, running from its temporary folder
    expect(script).toContain("npm-cli");                                // npm itself, installing globally ...
    expect(script).toMatch(/\(install\|i\|add\)/);
    expect(script).toContain("IndexOf($pkg");                           // ... this package
    expect(script).toContain(`$cur = ${process.pid}`);                  // the parents of this process ...
    expect(script).toContain("-not $ancestors.Contains($id)");          // ... are never taken for a rival
  });

  it("leaves npm's staging folders alone unless they are this package's", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-leftovers-"));
    dirs.push(dir);
    for (const name of [".excel-partner-AbC123", ".excel-partner-x9", ".excel-partner", ".other-pkg-AbC123", "excel-partner"]) fs.mkdirSync(path.join(dir, name));
    expect(removeNpmLeftovers(dir, "excel-partner")).toBe(2);
    expect(fs.readdirSync(dir).sort()).toEqual([".excel-partner", ".other-pkg-AbC123", "excel-partner"]);
    expect(removeNpmLeftovers(path.join(dir, "missing"), "excel-partner")).toBe(0);
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
