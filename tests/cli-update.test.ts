import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
// @ts-expect-error The npm executable ships as JavaScript, independently of the TS build.
import { readUpdateStatus, writeUpdateStatus } from "../bin/update.mjs";

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
