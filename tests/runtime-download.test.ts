import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { afterEach, describe, expect, it, vi } from "vitest";
// @ts-expect-error The updater's bin tools run independently of TypeScript.
import { installBinary } from "../bin/runtime-download.mjs";
// @ts-expect-error The updater's bin tools run independently of TypeScript.
import { installRuntime, runtimeInventory } from "../bin/runtime-manager.mjs";

const dirs: string[] = [];
const temporary = () => { const dir = fs.mkdtempSync(path.join(os.tmpdir(), "native-runtime-")); dirs.push(dir); return dir; };
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const bytes = Buffer.from("official release fixture");
const hash = crypto.createHash("sha512").update(bytes).digest("hex");
function item(release = {}) { return { bin: "agy", download: { [`${process.platform}-${process.arch}`]: { url: "https://example.test/agy", sha512: hash, ...release } } }; }
const target = (dir: string) => path.join(dir, process.platform === "win32" ? "agy.exe" : "agy");

describe("native CLI runtime installation", () => {
  it("checks the pin before installing a raw executable and removes staging data", async () => {
    const dir = temporary();
    await installBinary(dir, item(), { fetchRelease: async () => new Response(bytes) });
    expect(fs.readFileSync(target(dir))).toEqual(bytes);
    expect(fs.existsSync(path.join(dir, ".download"))).toBe(false);
  });
  it("never extracts or runs a payload whose checksum differs", async () => {
    const dir = temporary(), run = vi.fn();
    await expect(installBinary(dir, item({ archiveEntry: "antigravity" }), { run, fetchRelease: async () => new Response("tampered") })).rejects.toThrow("RUNTIME_INTEGRITY_FAILED");
    expect(run).not.toHaveBeenCalled();
    expect(fs.readdirSync(dir)).toEqual([]);
  });
  it("extracts only the pinned archive entry after verification", async () => {
    const dir = temporary(), run = vi.fn(async () => { fs.writeFileSync(path.join(dir, "antigravity"), "native executable"); return { code: 0 }; });
    await installBinary(dir, item({ archiveEntry: "antigravity" }), { run, fetchRelease: async () => new Response(bytes) });
    expect(run).toHaveBeenCalledWith("tar", ["-xzf", path.join(dir, ".download"), "-C", dir, "--", "antigravity"], expect.anything());
    expect(fs.readFileSync(target(dir), "utf8")).toBe("native executable");
  });
  it("rejects unsafe archive paths and unavailable platforms before downloading", async () => {
    const fetchRelease = vi.fn();
    await expect(installBinary(temporary(), item({ archiveEntry: "../agy" }), { fetchRelease })).rejects.toThrow("RUNTIME_MANIFEST_INVALID");
    await expect(installBinary(temporary(), { bin: "agy", download: {} }, { fetchRelease })).rejects.toThrow("RUNTIME_PLATFORM_UNSUPPORTED");
    expect(fetchRelease).not.toHaveBeenCalled();
  });
  it("uses the existing activation, reuse and version checks for binary providers", async () => {
    const dataDir = temporary();
    const download = vi.fn(async (dir: string) => fs.writeFileSync(target(dir), "native"));
    const run = vi.fn(async () => ({ code: 0, stdout: "1.5.16", stderr: "" }));
    const preflight = vi.fn(async () => { throw new Error("NATIVE_PREFLIGHT_AUTH_REQUIRED"); });
    const options = { dataDir, download, run, preflight, install: vi.fn() };
    const profile = await installRuntime(process.cwd(), "antigravity", options);
    expect(profile.command).toBe(target(path.join(dataDir, "runtimes", profile.id)));
    expect(runtimeInventory(dataDir).find((entry: any) => entry.provider === "antigravity")).toMatchObject({ managed: true, selected: true, version: "1.5.16", available: true });
    expect(await installRuntime(process.cwd(), "antigravity", options)).toEqual(profile);
    expect(download).toHaveBeenCalledOnce(); expect(preflight).toHaveBeenCalledOnce(); expect(options.install).not.toHaveBeenCalled();
  });
  it("does not activate a binary that reports the wrong version", async () => {
    const dataDir = temporary();
    await expect(installRuntime(process.cwd(), "antigravity", { dataDir,
      download: async (dir: string) => fs.writeFileSync(target(dir), "native"),
      run: async () => ({ code: 0, stdout: "9.0.0", stderr: "" }), preflight: vi.fn(),
    })).rejects.toThrow("CLI_PREFLIGHT_FAILED");
    expect(runtimeInventory(dataDir).find((entry: any) => entry.provider === "antigravity")).toMatchObject({ managed: false, selected: false, failure: "CLI_PREFLIGHT_FAILED" });
  });
});
