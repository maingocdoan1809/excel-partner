import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { UpdateError, UpdateService, compareVersions, fetchLatestFromRegistry, type UpdateServiceOptions } from "../bridge/src/updater.js";

describe("compareVersions", () => {
  it("orders versions by number, not by text", () => {
    expect(compareVersions("0.1.9", "0.1.19")).toBe(-1);
    expect(compareVersions("0.1.19", "0.1.9")).toBe(1);
    expect(compareVersions("0.1.19", "0.1.19")).toBe(0);
    expect(compareVersions("1.0.0", "0.99.99")).toBe(1);
    expect(compareVersions("0.2.0", "0.10.0")).toBe(-1);
  });
  it("puts a pre-release before its release", () => {
    expect(compareVersions("1.0.0-beta.1", "1.0.0")).toBe(-1);
    expect(compareVersions("1.0.0", "1.0.0-rc.1")).toBe(1);
    expect(compareVersions("1.0.0-alpha", "1.0.0-beta")).toBe(-1);
  });
  it("returns null for anything that is not a version", () => {
    expect(compareVersions("latest", "0.1.0")).toBeNull();
    expect(compareVersions("1.2", "1.2.3")).toBeNull();
    expect(compareVersions("", "1.2.3")).toBeNull();
  });
});

describe("fetchLatestFromRegistry", () => {
  const reply = (body: unknown, ok = true, status = 200) => vi.fn(async () => ({ ok, status, json: async () => body })) as unknown as typeof fetch;
  it("reads the version of the latest release", async () => {
    const fetchImpl = reply({ name: "excel-partner", version: "0.1.20" });
    expect(await fetchLatestFromRegistry(fetchImpl)).toBe("0.1.20");
    expect(vi.mocked(fetchImpl).mock.calls[0]![0]).toBe("https://registry.npmjs.org/excel-partner/latest");
  });
  it("fails clearly when the registry is down or answers nonsense", async () => {
    await expect(fetchLatestFromRegistry(reply({}, false, 503))).rejects.toThrow("HTTP 503");
    await expect(fetchLatestFromRegistry(reply({ version: "soon" }))).rejects.toThrow("no version");
    await expect(fetchLatestFromRegistry(reply({ version: 5 }))).rejects.toThrow("no version");
  });
});

describe("UpdateService", () => {
  let dir: string;
  let packageRoot: string;
  let dataDir: string;
  const binSource = path.resolve(__dirname, "..", "bin");
  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "updater-"));
    packageRoot = path.join(dir, "global", "excel-partner");
    dataDir = path.join(dir, "data");
    fs.mkdirSync(path.join(packageRoot, "bin"), { recursive: true });
    for (const file of ["update.mjs", "update-runner.mjs"]) fs.copyFileSync(path.join(binSource, file), path.join(packageRoot, "bin", file));
    fs.writeFileSync(path.join(packageRoot, "package.json"), JSON.stringify({ name: "excel-partner", version: "0.1.19" }));
  });
  afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

  const service = (extra: Partial<UpdateServiceOptions> = {}) =>
    new UpdateService({ packageRoot, dataDir, platform: "win32", globalRoot: () => path.join(dir, "global"), fetchLatest: async () => "0.1.20", launch: vi.fn(() => 4242), ...extra });
  const writeStatus = (status: object) => { fs.mkdirSync(dataDir, { recursive: true }); fs.writeFileSync(path.join(dataDir, "update-status.json"), JSON.stringify(status)); };

  it("reports the installed version from package.json and whether a newer one is out", async () => {
    const check = await service().check();
    expect(check).toMatchObject({ current: "0.1.19", latest: "0.1.20", updateAvailable: true, canUpdate: true, status: { state: "idle" } });
    expect(check.logFile).toBe(path.join(dataDir, "update.log"));
    expect(await service({ fetchLatest: async () => "0.1.19" }).check()).toMatchObject({ updateAvailable: false });
    expect(await service({ fetchLatest: async () => "0.1.3" }).check()).toMatchObject({ updateAvailable: false });
  });

  it("says why the latest version is unknown instead of failing", async () => {
    const check = await service({ fetchLatest: async () => { throw new Error("getaddrinfo ENOTFOUND registry.npmjs.org"); } }).check();
    expect(check).toMatchObject({ latest: null, updateAvailable: false, checkError: expect.stringContaining("ENOTFOUND") });
  });

  it("asks the registry once and then answers from memory until told to look again", async () => {
    const fetchLatest = vi.fn(async () => "0.1.20");
    const s = service({ fetchLatest });
    await s.check();
    await s.check();
    expect(fetchLatest).toHaveBeenCalledTimes(1);
    await s.check(true);
    expect(fetchLatest).toHaveBeenCalledTimes(2);
  });

  it("looks again after the cache runs out", async () => {
    let now = 1_000;
    const fetchLatest = vi.fn(async () => "0.1.20");
    const s = service({ fetchLatest, now: () => now, cacheMs: 5_000 });
    await s.check();
    now += 6_000;
    await s.check();
    expect(fetchLatest).toHaveBeenCalledTimes(2);
  });

  it("cannot replace a copy that is not the global npm install, or run off Windows", async () => {
    const elsewhere = await service({ globalRoot: () => path.join(dir, "somewhere else") }).check();
    expect(elsewhere).toMatchObject({ canUpdate: false, reason: expect.stringContaining("npm install -g") });
    expect(await service({ globalRoot: () => null }).check()).toMatchObject({ canUpdate: false, reason: expect.stringContaining("npm was not found") });
    expect(await service({ platform: "linux" }).check()).toMatchObject({ canUpdate: false, reason: expect.stringContaining("Windows") });
  });

  it("starts the update in a separate process from copies of the runner, with the settings it needs", () => {
    const launch = vi.fn(() => 4242);
    fs.mkdirSync(path.join(dir, "gw"), { recursive: true });
    fs.writeFileSync(path.join(dir, "gw", "token.txt"), "x");
    const status = service({ launch, gatewayDir: path.join(dir, "gw"), spec: "excel-partner@latest" }).start({ busy: false });
    expect(status).toMatchObject({ state: "running", from: "0.1.19" });
    const [runner, env, log] = launch.mock.calls[0] as unknown as [string, NodeJS.ProcessEnv, string];
    expect(path.basename(runner)).toBe("update-runner.mjs");
    expect(runner.startsWith(packageRoot)).toBe(false); // it must not run from the folder being replaced
    expect(fs.existsSync(path.join(path.dirname(runner), "update.mjs"))).toBe(true);
    expect(fs.readFileSync(runner, "utf8")).toBe(fs.readFileSync(path.join(packageRoot, "bin", "update-runner.mjs"), "utf8"));
    expect(env).toMatchObject({ EP_ROOT: packageRoot, EP_DATA: dataDir, EP_SPEC: "excel-partner@latest", EP_FROM: "0.1.19", EP_GATEWAY: "1" });
    expect(log).toBe(path.join(dataDir, "update.log"));
    expect(JSON.parse(fs.readFileSync(path.join(dataDir, "update-status.json"), "utf8"))).toMatchObject({ state: "running" });
    fs.rmSync(path.dirname(runner), { recursive: true, force: true });
  });

  it("does not start the gateway again when it was never set up", () => {
    const launch = vi.fn(() => 1);
    service({ launch, gatewayDir: path.join(dir, "no-gateway") }).start({ busy: false });
    expect((launch.mock.calls[0] as unknown as [string, NodeJS.ProcessEnv])[1].EP_GATEWAY).toBe("0");
    fs.rmSync(path.dirname((launch.mock.calls[0] as unknown as [string])[0]), { recursive: true, force: true });
  });

  it("refuses to start while the assistant is working, while another update runs, or when unsupported", () => {
    const launch = vi.fn(() => 1);
    expect(() => service({ launch }).start({ busy: true })).toThrow(expect.objectContaining({ code: "busy" }));
    expect(() => service({ launch, platform: "linux" }).start({ busy: false })).toThrow(expect.objectContaining({ code: "unsupported" }));
    writeStatus({ state: "running", pid: process.pid, startedAt: new Date().toISOString() });
    expect(() => service({ launch }).start({ busy: false })).toThrow(expect.objectContaining({ code: "running" }));
    expect(launch).not.toHaveBeenCalled();
  });

  it("records a failed start and says so when the runner cannot be launched", () => {
    const s = service({ launch: () => { throw new Error("spawn EPERM"); } });
    expect(() => s.start({ busy: false })).toThrow(UpdateError);
    expect(s.status()).toMatchObject({ state: "failed", error: "spawn EPERM" });
  });

  describe("status", () => {
    it("is idle before any update", () => expect(service().status()).toEqual({ state: "idle" }));

    it("passes on what the runner wrote, without its pid", () => {
      writeStatus({ state: "done", step: "done", from: "0.1.19", to: "0.1.20", pid: 1, startedAt: new Date().toISOString() });
      const status = service().status();
      expect(status).toMatchObject({ state: "done", to: "0.1.20" });
      expect(status).not.toHaveProperty("pid");
    });

    it("keeps saying running while the runner is alive", () => {
      writeStatus({ state: "running", step: "installing", message: "Installing the new version", pid: 77, startedAt: new Date().toISOString() });
      expect(service({ isPidAlive: () => true }).status()).toMatchObject({ state: "running", step: "installing" });
    });

    it("calls an update failed when its runner has disappeared", () => {
      writeStatus({ state: "running", pid: 77, startedAt: new Date().toISOString() });
      expect(service({ isPidAlive: () => false }).status()).toMatchObject({ state: "failed", error: expect.stringContaining("stopped") });
    });

    it("trusts a start that has no pid yet for a few seconds, then gives up on it", () => {
      const started = Date.parse("2026-11-04T10:00:00Z");
      writeStatus({ state: "running", startedAt: new Date(started).toISOString() });
      expect(service({ now: () => started + 5_000 }).status().state).toBe("running");
      expect(service({ now: () => started + 25_000 }).status().state).toBe("failed");
    });

    it("gives up on a record that has been running for far too long", () => {
      const started = Date.parse("2026-11-04T10:00:00Z");
      writeStatus({ state: "running", pid: 77, startedAt: new Date(started).toISOString() });
      expect(service({ isPidAlive: () => true, now: () => started + 16 * 60_000 }).status().state).toBe("failed");
    });

    it("ignores a damaged record", () => {
      fs.mkdirSync(dataDir, { recursive: true });
      fs.writeFileSync(path.join(dataDir, "update-status.json"), "{ not json");
      expect(service().status()).toEqual({ state: "idle" });
    });
  });
});
