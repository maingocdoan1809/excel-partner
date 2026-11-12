import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { installRuntime, runtimeInventory, verifyRuntime, repairRuntime } from "../bin/runtime-manager.mjs";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { snapshotRelease, updateRelease, verifyRelease, retireOldReleases } from "../bin/deployment.mjs";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { ensureBaseline, releaseOwners, verifiedActiveRelease, managementMain, installRecovery } from "../bin/manage.mjs";
// @ts-expect-error Standalone diagnostic tools.
import { collectDoctor } from "../bin/doctor-checks.mjs";
const store = createRequire(import.meta.url)(path.resolve("bin/runtime-store.cjs"));
/** Folders in a data directory, ignoring the `.trash-*` ones that are being deleted in the background. */
const folders = (directory: string) => fs.readdirSync(directory).filter(name => !name.startsWith(".trash-"));
const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function fixture() {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "managed-runtimes-")); dirs.push(directory);
  const root = path.join(directory, "source"), dataDir = path.join(directory, "owned");
  const put = (root: string, file: string, value: string) => { const target = path.join(root, file); fs.mkdirSync(path.dirname(target), { recursive: true }); fs.writeFileSync(target, value); return target; };
  const app = (root: string, version: string) => {
    put(root, "package.json", JSON.stringify({ name: "office-partner", version, type: "module" }));
    for (const file of ["bin/office-partner.mjs", "dist/bridge/src/index.js", "gateway/dist/server.mjs", "addin/dist/taskpane.js"]) put(root, file, "// fixture " + version);
    put(root, "bin/runtime-store.cjs", fs.readFileSync("bin/runtime-store.cjs", "utf8"));
    put(root, "bin/runtime-manifest.json", fs.readFileSync("bin/runtime-manifest.json", "utf8"));
    put(root, "node_modules/@github/copilot-sdk/package.json", JSON.stringify({ name: "@github/copilot-sdk", version: "1.0.18", main: "index.js" }));
    put(root, "node_modules/@github/copilot-sdk/index.js", "module.exports = {};");
  };
  app(root, "1.0.0");
  const install = vi.fn(async (prefix: string, spec: string) => {
    if (spec.startsWith("office-partner@")) { app(path.join(prefix, "node_modules", "office-partner"), spec.slice(15)); return; }
    const split = spec.lastIndexOf("@"), name = spec.slice(0, split), version = spec.slice(split + 1);
    const bin = ({ "@github/copilot": "copilot", "@anthropic-ai/claude-code": "claude", "@openai/codex": "codex", "opencode-ai": "opencode" } as Record<string, string>)[name]!;
    put(prefix, `node_modules/${name}/package.json`, JSON.stringify({ name, version, bin: { [bin]: `bin/${bin}.js` } }));
    put(prefix, `node_modules/${name}/bin/${bin}.js`, JSON.stringify({ version }));
  });
  const run = vi.fn(async (command: string) => ({ code: 0, stdout: JSON.parse(fs.readFileSync(command, "utf8")).version, stderr: "" }));
  const env = { PATH: "", Path: "" };
  const calls: string[] = [];
  const controls = { drain: vi.fn(async () => { calls.push("drain"); }), undrain: vi.fn(async () => { calls.push("undrain"); }),
    stop: vi.fn(async (release: { version: string }) => { calls.push("stop:" + release.version); }),
    start: vi.fn(async (release: { version: string }) => { calls.push("start:" + release.version); }) };
  const preflight = vi.fn(async () => {});
  return { root, dataDir, directory, put, app, install, run, env, calls, controls, preflight };
}

describe("managed runtime installation", () => {
  it.skipIf(process.platform !== "win32")("uses OpenCode's packaged native executable instead of the Windows batch launcher", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const install = async (prefix: string, spec: string) => {
      await w.install(prefix, spec);
      const version = spec.slice(spec.lastIndexOf("@") + 1), nativeName = `opencode-windows-${process.arch}`;
      w.put(prefix, "node_modules/opencode-ai/package.json", JSON.stringify({ name: "opencode-ai", version, bin: { opencode: "bin/opencode.cmd" } }));
      w.put(prefix, "node_modules/opencode-ai/bin/opencode.cmd", "@echo off\nstart opencode.exe");
      w.put(prefix, `node_modules/${nativeName}/package.json`, JSON.stringify({ name: nativeName, version }));
      w.put(prefix, `node_modules/${nativeName}/bin/opencode.exe`, JSON.stringify({ version }));
    };
    const profile = await installRuntime(w.root, "opencode", { ...w, install, adapterRoot: baseline.root });
    expect(profile.command).toContain(`opencode-windows-${process.arch}`);
    expect(path.extname(profile.command)).toBe(".exe");
    expect(store.childEnvironment({ OPENCODE_BIN_PATH: "user-global.exe" }, profile)).not.toHaveProperty("OPENCODE_BIN_PATH");
  });
  it("the standalone native preflight rejects unsigned-in providers before creating a session", () => {
    const w = fixture();
    w.put(w.root, "bin/provider-preflight.mjs", fs.readFileSync("bin/provider-preflight.mjs", "utf8"));
    w.put(w.root, "dist/providers/src/claude.js", `export class ClaudeAdapter {
      async detect() { return { installed: true, status: 'login_required' }; }
      async startSession() { throw new Error('START_SESSION_MUST_NOT_BE_CALLED'); }
      async dispose() {}
    }`);
    const result = spawnSync(process.execPath, [path.join(w.root, "bin/provider-preflight.mjs"), "claude", path.join(w.root, "claude.exe")], { encoding: "utf8", windowsHide: true, timeout: 5000 });
    expect(result.status).toBe(1);
    expect(result.stdout).toContain('"code":"AUTH_REQUIRED"');
    expect(result.stdout + result.stderr).not.toContain("START_SESSION_MUST_NOT_BE_CALLED");
  });
  it("installs the exact declared CLI in owned storage and keeps the adapter snapshot", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    expect(w.install.mock.calls[0]![1]).toBe("@github/copilot@1.0.95");
    expect(store.inside(path.join(w.dataDir, "runtimes"), profile.command)).toBe(true);
    expect(profile.adapterRoot).toBe(baseline.root);
    expect(runtimeInventory(w.dataDir).find((item: any) => item.provider === "copilot")).toMatchObject({ mode: "managed", version: "1.0.95", available: true });
    expect(w.env).not.toHaveProperty("COPILOT_AUTO_UPDATE");
  });

  it("does not change an active runtime when a staged download fails", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const old = store.readState(w.dataDir).active.copilot;
    const otherRoot = snapshotRelease(w.root, w.dataDir).root;
    await expect(installRuntime(w.root, "copilot", { ...w, adapterRoot: otherRoot, install: async () => { throw new Error("offline"); } })).rejects.toThrow("offline");
    expect(store.readState(w.dataDir).active.copilot).toBe(old);
    expect(fs.existsSync(profile.command)).toBe(true);
  });

  it("keeps the old selection when native authentication fails and activates only after verification", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const old = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const nextRoot = snapshotRelease(w.root, w.dataDir).root;
    await expect(installRuntime(w.root, "copilot", { ...w, adapterRoot: nextRoot, preflight: async () => { throw new Error("NATIVE_PREFLIGHT_AUTH_REQUIRED"); } })).rejects.toThrow("AUTH_REQUIRED");
    const state = store.readState(w.dataDir), pending = state.profiles[state.pending.copilot];
    expect(state.active.copilot).toBe(old.id);
    expect(runtimeInventory(w.dataDir)[1]).toMatchObject({ pendingVersion: "1.0.95", pendingVerified: false });
    expect(store.childEnvironment({ COPILOT_HOME: "user-global" }, pending)).toMatchObject({ COPILOT_HOME: pending.stateHome, COPILOT_CACHE_HOME: path.join(pending.stateHome, "cache"), COPILOT_AUTO_UPDATE: "false" });
    await installRuntime(w.root, "copilot", { ...w, adapterRoot: nextRoot });
    expect(store.readState(w.dataDir)).toMatchObject({ active: { copilot: pending.id }, previous: { copilot: old.id }, pending: {}, failures: {} });
  });

  it("detects a changed dependency even when the CLI entrypoint is intact, repairs it, and cleans quarantine on uninstall", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    w.put(path.join(w.dataDir, "runtimes", profile.id), "node_modules/@github/copilot/package.json", "{}");
    expect(() => verifyRuntime(profile, w.dataDir)).toThrow("RUNTIME_INTEGRITY_FAILED");
    const diagnostic = await collectDoctor({ root: w.root, env: { ...w.env, OFFICE_PARTNER_DATA_DIR: w.dataDir }, provider: "copilot", integrity: true, platform: "linux" });
    expect(diagnostic.checks).toContainEqual(expect.objectContaining({ status: "fail", code: "RUNTIME_INTEGRITY_FAILED" }));
    expect(JSON.stringify(diagnostic)).not.toContain(w.directory);
    const repaired = await repairRuntime(w.root, "copilot", w);
    expect(verifyRuntime(repaired, w.dataDir)).toBe(true);
    expect(fs.readdirSync(path.join(w.dataDir, "runtimes")).some(name => name.includes(".damaged-"))).toBe(true);
    releaseOwners(w.dataDir, { addin: true, gateway: true });
    expect(fs.readdirSync(path.join(w.dataDir, "runtimes"))).toEqual([]);
  });

  it("repairs an inactive saved-chat profile without changing the currently selected provider", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const old = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const next = await installRuntime(w.root, "copilot", { ...w, adapterRoot: snapshotRelease(w.root, w.dataDir).root });
    w.put(path.join(w.dataDir, "runtimes", old.id), "node_modules/@github/copilot/package.json", "{}");
    await repairRuntime(old.adapterRoot, "copilot", { ...w, profileId: old.id });
    expect(verifyRuntime(old, w.dataDir)).toBe(true);
    expect(store.readState(w.dataDir)).toMatchObject({ active: { copilot: next.id }, pending: {} });
  });

  it("pins native conversations and rejects damaged profiles rather than replaying or switching them", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    store.bindConversation(w.dataDir, "copilot", "native", profile.id);
    expect(store.selectProfile(w.dataDir, "copilot", "native").id).toBe(profile.id);
    w.put(path.dirname(profile.command), path.basename(profile.command), "changed");
    expect(() => store.selectProfile(w.dataDir, "copilot", "native")).toThrow("RESUME_RUNTIME_MISSING");
    expect(() => store.selectProfile(w.dataDir, "copilot", "old-without-profile")).toThrow("LEGACY_RESUME_PROFILE_UNKNOWN");
  });

  it("refuses stored path traversal and preserves a live installation lock", () => {
    const w = fixture();
    store.changeState(w.dataDir, (state: any) => { const id = "a".repeat(64); state.profiles[id] = { id, provider: "copilot", mode: "managed", command: path.join(w.directory, "external.exe"), adapterRoot: w.root }; });
    expect(() => store.readState(w.dataDir)).toThrow("RUNTIME_PATH_INVALID");
    expect(() => store.lock(w.dataDir, "operation", () => store.lock(w.dataDir, "operation", () => {}))).toThrow("OPERATION_LOCKED");
  });
});

describe("complete release update and rollback", () => {
  it("R07 reports a damaged active release for repair instead of falling back to an older one", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    expect(Object.keys(store.currentDeployment(w.dataDir).releases)).toEqual([updated.id]);
    expect(fs.existsSync(baseline.directory)).toBe(false);
    w.put(updated.root, "dist/bridge/src/index.js", "broken");
    expect(() => verifiedActiveRelease(w.dataDir)).toThrow("ACTIVE_RELEASE_DAMAGED_USE_RECOVERY");
    expect(store.currentDeployment(w.dataDir).active).toBe(updated.id);
  });

  it("R06 repairs a damaged release by installing the same version again, then keeps only the fresh one", async () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    w.put(updated.root, "dist/bridge/src/index.js", "broken");
    const installs = w.install.mock.calls.length;
    await managementMain(["repair"], { root: w.root, dataDir: w.dataDir, lifecycle: w.controls, report: () => {}, deploy: { install: w.install, run: w.run, env: w.env, preflight: w.preflight } });
    expect(w.install.mock.calls.slice(installs).map(call => call[1])).toEqual(["office-partner@1.1.0"]);
    const deployment = store.currentDeployment(w.dataDir);
    expect(Object.keys(deployment.releases)).toHaveLength(1);
    expect(deployment.active).not.toBe(updated.id);
    expect(fs.existsSync(updated.directory)).toBe(false);
    expect(verifiedActiveRelease(w.dataDir)!.version).toBe("1.1.0");
  });

  it("R10 rollback means installing that version from npm, and needs a version", async () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    await expect(managementMain(["rollback"], { root: w.root, dataDir: w.dataDir, lifecycle: w.controls })).rejects.toThrow("USE_ROLLBACK_WITH_A_VERSION");
    await expect(managementMain(["rollback", "latest"], { root: w.root, dataDir: w.dataDir, lifecycle: w.controls })).rejects.toThrow("USE_ROLLBACK_WITH_A_VERSION");
    const installs = w.install.mock.calls.length;
    await managementMain(["rollback", "1.0.5"], { root: w.root, dataDir: w.dataDir, lifecycle: w.controls, report: () => {}, deploy: { install: w.install, run: w.run, env: w.env, preflight: w.preflight } });
    expect(w.install.mock.calls.slice(installs).map(call => call[1])).toEqual(["office-partner@1.0.5"]);
    const deployment = store.currentDeployment(w.dataDir);
    expect(deployment.releases[deployment.active].version).toBe("1.0.5");
    expect(Object.keys(deployment.releases)).toHaveLength(1);
  });
  it("R09 releases the final owner when gateway was uninstalled first and its config remains", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    store.changeState(w.dataDir, (state: any) => { state.owners = { addin: true, gateway: true }; });
    releaseOwners(w.dataDir, { addin: false, gateway: true, gatewayConfigured: true });
    expect(fs.existsSync(baseline.root)).toBe(true);
    releaseOwners(w.dataDir, { addin: true, gateway: false, gatewayConfigured: true });
    expect(fs.existsSync(baseline.root)).toBe(false);
    expect(store.currentDeployment(w.dataDir).active).toBeNull();
  });

  it("always exits maintenance if repair encounters a busy second service", async () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    const controls = { ...w.controls, drain: async () => { throw new Error("SERVICE_BUSY"); } };
    await expect(managementMain(["repair"], { root: w.root, dataDir: w.dataDir, lifecycle: controls })).rejects.toThrow("SERVICE_BUSY");
    expect(w.controls.undrain).toHaveBeenCalledOnce();
    expect(w.controls.stop).not.toHaveBeenCalled(); expect(w.controls.start).not.toHaveBeenCalled();
  });

  it("copies an immutable recovery toolset before replacing its launcher", () => {
    const w = fixture();
    w.put(w.root, "bin/manage.mjs", "// recovery fixture");
    installRecovery(w.root, w.dataDir);
    const launcher = fs.readFileSync(path.join(w.dataDir, "recovery/recovery.mjs"), "utf8");
    installRecovery(w.root, w.dataDir);
    expect(fs.readFileSync(path.join(w.dataDir, "recovery/recovery.mjs"), "utf8")).toBe(launcher);
    w.put(w.root, "bin/manage.mjs", "// next recovery fixture"); installRecovery(w.root, w.dataDir);
    expect(fs.readdirSync(path.join(w.dataDir, "recovery/versions"))).toHaveLength(2);
  });

  it("runs real standalone recovery management after the original package is removed", () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    installRecovery(path.resolve("."), w.dataDir);
    expect(store.inside(w.directory, w.root)).toBe(true); fs.rmSync(w.root, { recursive: true, force: true });
    const result = spawnSync(process.execPath, [path.join(w.dataDir, "recovery/recovery.mjs"), "versions"], {
      env: { ...process.env, PATH: "", Path: "" }, windowsHide: true, encoding: "utf8", timeout: 5000,
    });
    expect(result.status, result.stderr).toBe(0);
    expect(JSON.parse(result.stdout)).toMatchObject({ active: "1.0.0", providers: expect.any(Array) });
  });
  it("prepares release and runtimes before stopping, then keeps only the new release and its CLI", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    store.bindConversation(w.dataDir, "copilot", "old-chat", original.id);
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    expect(w.calls).toEqual(["drain", "stop:1.0.0", "start:1.1.0"]);
    const deployment = store.currentDeployment(w.dataDir), state = store.readState(w.dataDir);
    expect(deployment).toMatchObject({ active: updated.id });
    expect(Object.keys(deployment.releases)).toEqual([updated.id]);
    expect(folders(path.join(w.dataDir, "releases"))).toEqual([updated.id]);
    // Same CLI version: the old profile only duplicated the new one, so it and its folder are gone and the chat moved over.
    expect(state.profiles[original.id]).toBeUndefined();
    expect(fs.existsSync(path.join(w.dataDir, "runtimes", original.id))).toBe(false);
    // The old folders were renamed at once and are deleted in the background.
    await vi.waitFor(() => { expect(fs.readdirSync(path.join(w.dataDir, "releases"))).toEqual([updated.id]); expect(fs.readdirSync(path.join(w.dataDir, "runtimes")).filter(name => name.startsWith(".trash-"))).toEqual([]); });
    expect(state.conversations[store.conversationKey("copilot", "old-chat")]).toBe(state.active.copilot);
    expect(store.selectProfile(w.dataDir, "copilot", "old-chat").adapterRoot).toBe(updated.root);
    expect(state.previous.copilot).toBeUndefined();
  });

  it("R11 keeps an older CLI that saved chats still resume with, served by the current release, and nothing else", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    store.bindConversation(w.dataDir, "copilot", "old-chat", original.id);
    // The next release ships a different Copilot CLI.
    const next = JSON.parse(fs.readFileSync("bin/runtime-manifest.json", "utf8")); next.providers.copilot.version = "9.9.9";
    const install = async (prefix: string, spec: string) => { await w.install(prefix, spec); if (spec.startsWith("office-partner@")) w.put(path.join(prefix, "node_modules", "office-partner"), "bin/runtime-manifest.json", JSON.stringify(next)); };
    const run = vi.fn(async (command: string) => ({ code: 0, stdout: JSON.parse(fs.readFileSync(command, "utf8")).version, stderr: "" }));
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, install, run, ...w.controls });
    const state = store.readState(w.dataDir);
    expect(state.profiles[original.id]).toMatchObject({ version: original.version, adapterRoot: updated.root });
    expect(fs.existsSync(original.command)).toBe(true);
    expect(state.profiles[state.active.copilot].version).toBe("9.9.9");
    expect(store.selectProfile(w.dataDir, "copilot", "old-chat").id).toBe(original.id);
    expect(Object.keys(store.currentDeployment(w.dataDir).releases)).toEqual([updated.id]);
    expect(folders(path.join(w.dataDir, "releases"))).toEqual([updated.id]);
  });

  it("R12 drops runtimes that nothing selects, uses or has chats on, and sweeps release folders left behind earlier", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const stray = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root, activate: false, trackPending: false });
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    expect(store.readState(w.dataDir).profiles[stray.id]).toBeUndefined();
    const orphan = path.join(w.dataDir, "releases", "a".repeat(64)); fs.mkdirSync(orphan, { recursive: true });
    const staging = path.join(w.dataDir, "releases", ".install-in-progress"); fs.mkdirSync(staging, { recursive: true });
    const swept = retireOldReleases(w.dataDir);
    expect(swept.releases).toEqual([]);
    await swept.pending;
    expect(fs.existsSync(orphan)).toBe(false); expect(fs.existsSync(staging)).toBe(true);
    expect(store.currentDeployment(w.dataDir).active).toBe(updated.id);
  });

  it("R13 never fails because an old folder cannot be renamed or deleted, and the next update finishes the job", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const realRename = fs.renameSync.bind(fs);
    const rename = vi.spyOn(fs, "renameSync").mockImplementation((from, to) => {
      if (String(to).includes(".trash-")) throw Object.assign(new Error("EBUSY"), { code: "EBUSY" });
      return realRename(from, to);
    });
    const remove = vi.spyOn(fs.promises, "rm").mockRejectedValue(Object.assign(new Error("EBUSY"), { code: "EBUSY" }));
    try {
      const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
      expect(store.currentDeployment(w.dataDir)).toMatchObject({ active: updated.id });
    } finally { rename.mockRestore(); remove.mockRestore(); }
    expect(fs.existsSync(baseline.directory)).toBe(true);
    await retireOldReleases(w.dataDir).pending;
    expect(fs.existsSync(baseline.directory)).toBe(false);
  });

  it("keeps the current release running if package installation fails or the service is busy", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    await expect(updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls, install: async () => { throw new Error("download failed"); } })).rejects.toThrow("download failed");
    expect(w.controls.stop).not.toHaveBeenCalled();
    await expect(updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls, drain: async () => { throw new Error("SERVICE_BUSY"); } })).rejects.toThrow("SERVICE_BUSY");
    expect(w.controls.stop).not.toHaveBeenCalled();
    expect(store.currentDeployment(w.dataDir).active).toBe(baseline.id);
  });

  it("reports a failed activation and does not go back to the older release by itself", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const start = vi.fn(async (release: { version: string }) => { w.calls.push("start:" + release.version); throw new Error("startup failed"); });
    await expect(updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls, start })).rejects.toThrow("startup failed");
    // The half-started services are closed; the old release is not started again.
    expect(w.calls).toEqual(["drain", "stop:1.0.0", "start:1.1.0", "stop:1.1.0", "undrain"]);
    const deployment = store.currentDeployment(w.dataDir);
    expect(deployment.active).not.toBe(baseline.id);
    expect(deployment.releases[deployment.active].version).toBe("1.1.0");
  });

  it("rejects corrupted release assets", () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    w.put(baseline.root, "dist/bridge/src/index.js", "corrupted");
    expect(() => verifyRelease(baseline)).toThrow("RELEASE_INTEGRITY_FAILED");
  });

  it("uninstall removes only owned artifacts when all owners are gone, preserving user data and external CLIs", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const database = w.put(w.dataDir, "history.db", "preserve user database");
    const history = w.put(w.dataDir, "usage.jsonl", "keep"), external = w.put(w.directory, "user-cli.exe", "keep");
    store.changeState(w.dataDir, (state: any) => { state.owners = { addin: true, gateway: true }; });
    releaseOwners(w.dataDir, { addin: true, gateway: false });
    expect(fs.existsSync(profile.command)).toBe(true);
    releaseOwners(w.dataDir, { addin: false, gateway: true });
    expect(fs.existsSync(profile.command)).toBe(false);
    expect(fs.existsSync(baseline.root)).toBe(false);
    expect(fs.readFileSync(database, "utf8")).toBe("preserve user database");
    expect(fs.readFileSync(history, "utf8")).toBe("keep"); expect(fs.readFileSync(external, "utf8")).toBe("keep");
  });
});
