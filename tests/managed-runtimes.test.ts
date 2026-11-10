import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { installRuntime, runtimeInventory, verifyRuntime, repairRuntime } from "../bin/runtime-manager.mjs";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { snapshotRelease, updateRelease, rollbackRelease, verifyRelease, recoverInterrupted } from "../bin/deployment.mjs";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { ensureBaseline, releaseOwners, verifiedActiveRelease, managementMain, installRecovery } from "../bin/manage.mjs";
// @ts-expect-error Standalone diagnostic tools.
import { collectDoctor } from "../bin/doctor-checks.mjs";
const store = createRequire(import.meta.url)(path.resolve("bin/runtime-store.cjs"));
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
  it("restores the retained release's provider selection when startup detects corrupt active files", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    store.bindConversation(w.dataDir, "copilot", "new-chat", store.readState(w.dataDir).active.copilot);
    w.put(updated.root, "dist/bridge/src/index.js", "broken");
    expect(verifiedActiveRelease(w.dataDir).id).toBe(baseline.id);
    expect(store.readState(w.dataDir).active.copilot).toBe(original.id);
    expect(store.readState(w.dataDir).conversations[store.conversationKey("copilot", "new-chat")]).toBeTruthy();
    expect(store.currentDeployment(w.dataDir)).toMatchObject({ held: true, previous: null, journal: { phase: "integrity_fallback" } });
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
  it("prepares release and runtimes before stopping, and rolls back from cached files offline", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    expect(w.calls).toEqual(["drain", "stop:1.0.0", "start:1.1.0"]);
    expect(store.currentDeployment(w.dataDir).active).toBe(updated.id);
    expect(store.readState(w.dataDir).active.copilot).not.toBe(original.id);
    store.bindConversation(w.dataDir, "copilot", "old-chat", original.id);
    const before = w.install.mock.calls.length;
    await rollbackRelease(undefined, { dataDir: w.dataDir, ...w.controls });
    expect(w.install.mock.calls).toHaveLength(before);
    expect(store.currentDeployment(w.dataDir)).toMatchObject({ active: baseline.id, held: true });
    expect(store.readState(w.dataDir).active.copilot).toBe(original.id);
    expect(store.selectProfile(w.dataDir, "copilot", "old-chat").id).toBe(original.id);
  });

  it("keeps the current release running if package installation fails or the service is busy", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    await expect(updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls, install: async () => { throw new Error("download failed"); } })).rejects.toThrow("download failed");
    expect(w.controls.stop).not.toHaveBeenCalled();
    await expect(updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls, drain: async () => { throw new Error("SERVICE_BUSY"); } })).rejects.toThrow("SERVICE_BUSY");
    expect(w.controls.stop).not.toHaveBeenCalled();
    expect(store.currentDeployment(w.dataDir).active).toBe(baseline.id);
  });

  it("restores both release and runtime selection after activation failure", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const start = vi.fn(async (release: { version: string }) => { w.calls.push("start:" + release.version); if (release.version === "1.1.0") throw new Error("startup failed"); });
    await expect(updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls, start })).rejects.toThrow("ACTIVATION_FAILED_PREVIOUS_RELEASE_RESTORED");
    expect(w.calls).toEqual(["drain", "stop:1.0.0", "start:1.1.0", "stop:1.1.0", "start:1.0.0"]);
    expect(store.currentDeployment(w.dataDir).active).toBe(baseline.id);
    expect(store.readState(w.dataDir).active.copilot).toBe(original.id);
  });

  it("rejects corrupted cached release assets and recovers an interrupted activation", () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const other = snapshotRelease(w.root, w.dataDir);
    const deployment = store.currentDeployment(w.dataDir);
    deployment.active = other.id; deployment.journal = { phase: "activating", from: baseline.id, to: other.id, previous: null, runtimeBefore: store.readState(w.dataDir) };
    store.atomic(path.join(w.dataDir, "deployment.json"), deployment);
    recoverInterrupted(w.dataDir);
    expect(store.currentDeployment(w.dataDir)).toMatchObject({ active: baseline.id, held: true, journal: { phase: "recovered" } });
    w.put(other.root, "dist/bridge/src/index.js", "corrupted");
    expect(() => verifyRelease(other)).toThrow("RELEASE_INTEGRITY_FAILED");
  });

  it("uninstall removes only owned artifacts when all owners are gone, preserving user data and external CLIs", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", { ...w, adapterRoot: baseline.root });
    const history = w.put(w.dataDir, "usage.jsonl", "keep"), external = w.put(w.directory, "user-cli.exe", "keep");
    store.changeState(w.dataDir, (state: any) => { state.owners = { addin: true, gateway: true }; });
    releaseOwners(w.dataDir, { addin: true, gateway: false });
    expect(fs.existsSync(profile.command)).toBe(true);
    releaseOwners(w.dataDir, { addin: false, gateway: true });
    expect(fs.existsSync(profile.command)).toBe(false);
    expect(fs.existsSync(baseline.root)).toBe(false);
    expect(fs.readFileSync(history, "utf8")).toBe("keep"); expect(fs.readFileSync(external, "utf8")).toBe("keep");
  });
});
