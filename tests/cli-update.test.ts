import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
// @ts-expect-error The npm executable ships as JavaScript, independently of the TS build.
import { updatePackage } from "../bin/update.mjs";

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
    expect(run.mock.calls[1]).toEqual(["npm.cmd", ["install", "--global", "excel-partner@latest"], expect.objectContaining({ cwd: opts.dataDir, shell: true })]);
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
});
