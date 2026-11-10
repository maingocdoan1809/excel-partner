import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { PLAYWRIGHT_CORE_VERSION, installedPlaywrightEntry, loadPlaywright } from "../bridge/src/playwright-loader.js";

const missing = () => Object.assign(new Error("Cannot find package 'playwright-core'"), { code: "ERR_MODULE_NOT_FOUND" });
const driver = { chromium: { launch: vi.fn() } };

let tools = "";
beforeEach(() => { tools = fs.mkdtempSync(path.join(os.tmpdir(), "ep-tools-")); });
afterEach(() => fs.rmSync(tools, { recursive: true, force: true }));

describe("playwright-core on demand", () => {
  it("uses the copy a source checkout already has, without touching npm", async () => {
    const run = vi.fn();
    const out = vi.fn();
    const loaded = await loadPlaywright(tools, out, { load: async () => driver, run });
    expect(loaded).toBe(driver);
    expect(run).not.toHaveBeenCalled();
    expect(out).not.toHaveBeenCalled();
  });

  it("downloads the pinned version once into the tools folder and loads it from there", async () => {
    const entry = installedPlaywrightEntry(tools);
    let installed = false;
    const run = vi.fn(() => { installed = true; return { status: 0, error: undefined, stderr: "" }; });
    const load = vi.fn(async (specifier: string) => {
      if (specifier === "playwright-core") throw missing();
      expect(specifier).toBe(pathToFileURL(entry).href);
      return driver;
    });
    const out = vi.fn();
    const loaded = await loadPlaywright(tools, out, { load, run, exists: () => installed, platform: "win32" });
    expect(loaded).toBe(driver);
    expect(run).toHaveBeenCalledTimes(1);
    const [command, args, options] = run.mock.calls[0] as unknown as [string, string[], { cwd: string; shell: boolean }];
    expect(command).toBe("npm.cmd");
    expect(args).toEqual(["install", "--no-audit", "--no-fund", "--no-package-lock", "--omit=dev", `playwright-core@${PLAYWRIGHT_CORE_VERSION}`]);
    expect(options).toMatchObject({ cwd: tools, shell: true });
    expect(JSON.parse(fs.readFileSync(path.join(tools, "package.json"), "utf8"))).toEqual({ name: "office-partner-tools", private: true });
    expect(out).toHaveBeenCalledWith(expect.stringContaining("Downloading the browser driver"));
  });

  it("reuses a downloaded copy without running npm again", async () => {
    const run = vi.fn();
    const load = vi.fn(async (specifier: string) => { if (specifier === "playwright-core") throw missing(); return driver; });
    await loadPlaywright(tools, vi.fn(), { load, run, exists: () => true, platform: "linux" });
    expect(run).not.toHaveBeenCalled();
    expect(load).toHaveBeenLastCalledWith(pathToFileURL(installedPlaywrightEntry(tools)).href);
  });

  it("says why when npm cannot download it", async () => {
    const load = vi.fn(async () => { throw missing(); });
    const run = vi.fn(() => ({ status: 1, error: undefined, stderr: "npm ERR! network\nnpm ERR! ETIMEDOUT" }));
    await expect(loadPlaywright(tools, vi.fn(), { load, run, exists: () => false, platform: "linux" }))
      .rejects.toThrow("npm could not download the browser driver (exit 1). npm ERR! ETIMEDOUT");
    const quiet = vi.fn(() => ({ status: 0, error: undefined, stderr: "" }));
    await expect(loadPlaywright(tools, vi.fn(), { load, run: quiet, exists: () => false, platform: "linux" }))
      .rejects.toThrow("did not install");
  });

  it("does not hide an unrelated import failure behind a download", async () => {
    const run = vi.fn();
    const load = vi.fn(async () => { throw new Error("SyntaxError inside the driver"); });
    await expect(loadPlaywright(tools, vi.fn(), { load, run })).rejects.toThrow("SyntaxError inside the driver");
    expect(run).not.toHaveBeenCalled();
  });
});
