import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { installRuntime, runtimeInventory, verifyRuntime, repairRuntime, uninstallRuntime } from "../bin/runtime-manager.mjs";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { updateRelease, verifyRelease, retireOldReleases } from "../bin/deployment.mjs";
// @ts-expect-error The management tools are published as standalone JavaScript.
import { ensureBaseline, releaseOwners, verifiedActiveRelease, managementMain, installRecovery } from "../bin/manage.mjs";
// @ts-expect-error Standalone diagnostic tools.
import { collectDoctor } from "../bin/doctor-checks.mjs";
const store = createRequire(import.meta.url)(path.resolve("bin/runtime-store.cjs"));
/** Folders in a data directory, ignoring the `.trash-*` ones that are being deleted in the background. */
const folders = (directory: string) => fs.readdirSync(directory).filter(name => !name.startsWith(".trash-"));
const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
const manifest = () => JSON.parse(fs.readFileSync("bin/runtime-manifest.json", "utf8"));
function fixture() {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "managed-runtimes-")); dirs.push(directory);
  const root = path.join(directory, "source"), dataDir = path.join(directory, "owned");
  const put = (root: string, file: string, value: string) => { const target = path.join(root, file); fs.mkdirSync(path.dirname(target), { recursive: true }); fs.writeFileSync(target, value); return target; };
  const app = (root: string, version: string) => {
    put(root, "package.json", JSON.stringify({ name: "office-partner", version, type: "module" }));
    for (const file of ["bin/office-partner.mjs", "dist/bridge/src/index.js", "gateway/dist/server.mjs", "frontend/dist/taskpane.js"]) put(root, file, "// fixture " + version);
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
  /** A product folder whose manifest names another Copilot CLI version, as the next release would. */
  const shipping = (copilot: string) => {
    const next = path.join(directory, "ships-" + copilot); app(next, "1.0.0");
    const changed = manifest(); changed.providers.copilot.version = copilot;
    put(next, "bin/runtime-manifest.json", JSON.stringify(changed));
    return next;
  };
  return { root, dataDir, directory, put, app, install, run, env, calls, controls, preflight, shipping };
}

describe("managed runtime installation", () => {
  it.skipIf(process.platform !== "win32")("uses OpenCode's packaged native executable instead of the Windows batch launcher", async () => {
    const w = fixture();
    const install = async (prefix: string, spec: string) => {
      await w.install(prefix, spec);
      const version = spec.slice(spec.lastIndexOf("@") + 1), nativeName = `opencode-windows-${process.arch}`;
      w.put(prefix, "node_modules/opencode-ai/package.json", JSON.stringify({ name: "opencode-ai", version, bin: { opencode: "bin/opencode.cmd" } }));
      w.put(prefix, "node_modules/opencode-ai/bin/opencode.cmd", "@echo off\nstart opencode.exe");
      w.put(prefix, `node_modules/${nativeName}/package.json`, JSON.stringify({ name: nativeName, version }));
      w.put(prefix, `node_modules/${nativeName}/bin/opencode.exe`, JSON.stringify({ version }));
    };
    const profile = await installRuntime(w.root, "opencode", { ...w, install });
    expect(profile.command).toContain(`opencode-windows-${process.arch}`);
    expect(path.extname(profile.command)).toBe(".exe");
    expect(store.childEnvironment({ OPENCODE_BIN_PATH: "user-global.exe" }, profile)).not.toHaveProperty("OPENCODE_BIN_PATH");
  });

  it("keeps the CLI a provider already runs when a release changes anything else in the manifest, and downloads nothing", async () => {
    const w = fixture();
    const first = await installRuntime(w.root, "copilot", { ...w });
    const downloads = () => w.install.mock.calls.filter(([, spec]) => String(spec).startsWith("@github/copilot@")).length;
    expect(downloads()).toBe(1);
    store.changeState(w.dataDir, (state: { active: Record<string, string> }) => { state.active.copilot = first.id; });
    const next = path.join(w.directory, "ships-same-cli"); w.app(next, "1.0.0");
    const changed = manifest();
    changed.providers.copilot.minVersion = "1.0.0";
    changed.providers.copilot.env = { ...changed.providers.copilot.env, state: { ...changed.providers.copilot.env.state, SOMETHING_NEW: "." } };
    w.put(next, "bin/runtime-manifest.json", JSON.stringify(changed));
    const [checksBefore, preflightsBefore] = [w.run.mock.calls.length, w.preflight.mock.calls.length];
    const second = await installRuntime(next, "copilot", { ...w });
    expect(second.id).toBe(first.id);
    expect(downloads()).toBe(1);
    // Nothing about the CLI is run or checked again: it was checked when it was installed.
    expect(w.run.mock.calls.length).toBe(checksBefore);
    expect(w.preflight.mock.calls.length).toBe(preflightsBefore);
  });

  it("the standalone native preflight reports a provider that is not signed in with a reason code", () => {
    const w = fixture();
    w.put(w.root, "bin/provider-preflight.mjs", fs.readFileSync("bin/provider-preflight.mjs", "utf8"));
    w.put(w.root, "dist/providers/src/registry.js", `export function createAdapter() {
      return { async verifyReady() { throw new Error("authentication required"); }, async dispose() {} };
    }`);
    const result = spawnSync(process.execPath, [path.join(w.root, "bin/provider-preflight.mjs"), "claude", path.join(w.root, "claude.exe")], { encoding: "utf8", windowsHide: true, timeout: 5000 });
    expect(result.status).toBe(1);
    expect(result.stdout).toContain('"code":"AUTH_REQUIRED"');
  });

  it("installs the exact declared CLI in owned storage and runs it from then on", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", w);
    expect(w.install.mock.calls[0]![1]).toBe("@github/copilot@1.0.95");
    expect(store.inside(path.join(w.dataDir, "runtimes"), profile.command)).toBe(true);
    expect(store.readState(w.dataDir)).toMatchObject({ selected: ["copilot"], active: { copilot: profile.id } });
    expect(store.selectProfile(w.dataDir, "copilot").id).toBe(profile.id);
    expect(runtimeInventory(w.dataDir).find((item: any) => item.provider === "copilot")).toMatchObject({ managed: true, version: "1.0.95", available: true });
    expect(w.env).not.toHaveProperty("COPILOT_AUTO_UPDATE");
  });

  it("uninstalling stops a provider from using its CLI, and installing again turns it back on", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", w);
    uninstallRuntime("copilot", w);
    expect(store.readState(w.dataDir)).toMatchObject({ selected: [], active: {}, disabled: ["copilot"], profiles: {} });
    expect(() => store.selectProfile(w.dataDir, "copilot")).toThrow("PROVIDER_DISABLED");
    await installRuntime(w.root, "copilot", w);
    expect(store.readState(w.dataDir)).toMatchObject({ selected: ["copilot"], active: { copilot: profile.id }, disabled: [] });
    expect(() => uninstallRuntime("nope", w)).toThrow("PROVIDER_INVALID");
  });

  it("runs the system CLI for a provider the product installed nothing for", () => {
    const w = fixture();
    expect(store.selectProfile(w.dataDir, "codex")).toBeNull();
    expect(store.childEnvironment({ PATH: "x" }, null)).toMatchObject({ PATH: "x", COPILOT_AUTO_UPDATE: "false" });
  });

  it("gives each CLI its own state folder and only the settings it needs", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", w);
    expect(store.childEnvironment({ COPILOT_HOME: "user-global", COPILOT_CLI_PATH: "user-global" }, profile)).toMatchObject({
      COPILOT_HOME: profile.stateHome, COPILOT_CACHE_HOME: path.join(profile.stateHome, "cache"), COPILOT_BIN: profile.command, COPILOT_AUTO_UPDATE: "false" });
    expect(store.childEnvironment({ COPILOT_CLI_PATH: "user-global" }, profile)).not.toHaveProperty("COPILOT_CLI_PATH");
  });

  it("does not download a CLI it already has, and keeps its sign-in and settings when the version changes", async () => {
    const w = fixture();
    const first = await installRuntime(w.root, "copilot", w);
    const downloads = () => w.install.mock.calls.length;
    const again = await installRuntime(w.shipping("1.0.95"), "copilot", w);
    expect(again.id).toBe(first.id);
    expect(downloads()).toBe(1);
    const next = await installRuntime(w.shipping("9.9.9"), "copilot", w);
    expect(next.id).not.toBe(first.id);
    expect(downloads()).toBe(2);
    expect(next.stateHome).toBe(first.stateHome);
    expect(store.readState(w.dataDir)).toMatchObject({ active: { copilot: next.id } });
    expect(store.readState(w.dataDir).profiles[first.id]).toBeUndefined();
  });

  it("does not change what a provider runs when a download or a check of the new CLI fails", async () => {
    const w = fixture();
    const first = await installRuntime(w.root, "copilot", w);
    await expect(installRuntime(w.shipping("9.9.9"), "copilot", { ...w, install: async () => { throw new Error("offline"); } })).rejects.toThrow("offline");
    await expect(installRuntime(w.shipping("9.9.8"), "copilot", { ...w, preflight: async () => { throw new Error("NATIVE_PREFLIGHT_POLICY_DENIED"); } })).rejects.toThrow("POLICY_DENIED");
    const state = store.readState(w.dataDir);
    expect(state.active.copilot).toBe(first.id);
    expect(state.failures.copilot).toMatchObject({ code: "NATIVE_PREFLIGHT_POLICY_DENIED", version: "9.9.8" });
    expect(fs.existsSync(first.command)).toBe(true);
  });

  it("installs and uses a CLI whose person has not signed in yet, since that cannot be checked before", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", { ...w, preflight: async () => { throw new Error("NATIVE_PREFLIGHT_AUTH_REQUIRED"); } });
    expect(store.readState(w.dataDir)).toMatchObject({ active: { copilot: profile.id }, failures: {} });
  });

  it("detects a changed dependency even when the CLI entrypoint is intact, repairs it, and cleans quarantine on uninstall", async () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", w);
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

  it("refuses a CLI file that is missing or was changed, and a disabled provider, rather than running something else", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", w);
    w.put(path.dirname(profile.command), path.basename(profile.command), "changed");
    expect(() => store.selectProfile(w.dataDir, "copilot")).toThrow("RUNTIME_MISSING");
    store.changeState(w.dataDir, (state: any) => { state.disabled = ["copilot"]; });
    expect(() => store.selectProfile(w.dataDir, "copilot")).toThrow("PROVIDER_DISABLED");
  });

  it("checks the CLI file again only when it changes, not on every lookup", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", w);
    const hashed = vi.spyOn(fs, "readSync");
    store.selectProfile(w.dataDir, "copilot");
    const afterFirst = hashed.mock.calls.length;
    for (let i = 0; i < 5; i++) store.selectProfile(w.dataDir, "copilot");
    expect(hashed.mock.calls.length).toBe(afterFirst);
    hashed.mockRestore();
    expect(profile.fingerprint).toBeTruthy();
  });

  it("reads a state file written by an older version, keeping only the installs in use", async () => {
    const w = fixture();
    const profile = await installRuntime(w.root, "copilot", w);
    const old = store.readState(w.dataDir);
    fs.writeFileSync(path.join(w.dataDir, "runtime-state.json"), JSON.stringify({
      ...old, previous: { copilot: "a".repeat(64) }, legacy: { copilot: "b".repeat(64) }, pending: { copilot: "c".repeat(64) }, conversations: { "copilot:x": "a".repeat(64) },
      profiles: { ...old.profiles, ["b".repeat(64)]: { id: "b".repeat(64), provider: "copilot", mode: "external", command: path.join(w.directory, "system-cli.exe"), adapterRoot: w.root } },
    }));
    const state = store.readState(w.dataDir);
    expect(Object.keys(state.profiles)).toEqual([profile.id]);
    for (const dropped of ["previous", "legacy", "pending", "conversations"]) expect(state).not.toHaveProperty(dropped);
  });

  it("refuses stored path traversal and preserves a live installation lock", () => {
    const w = fixture();
    store.changeState(w.dataDir, (state: any) => { const id = "a".repeat(64); state.profiles[id] = { id, provider: "copilot", command: path.join(w.directory, "external.exe") }; });
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

  it("uninstall removes only owned artifacts when all owners are gone, preserving user data and external CLIs", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const profile = await installRuntime(w.root, "copilot", w);
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

  it("prepares release and CLI before stopping, and a CLI that has not changed is not installed again", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", w);
    const downloads = w.install.mock.calls.length;
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    expect(w.calls).toEqual(["drain", "stop:1.0.0", "start:1.1.0"]);
    expect(w.install.mock.calls.slice(downloads).map(call => call[1])).toEqual(["office-partner@1.1.0"]);
    const deployment = store.currentDeployment(w.dataDir), state = store.readState(w.dataDir);
    expect(deployment).toMatchObject({ active: updated.id });
    expect(Object.keys(deployment.releases)).toEqual([updated.id]);
    expect(folders(path.join(w.dataDir, "releases"))).toEqual([updated.id]);
    expect(state.active.copilot).toBe(original.id);
    expect(folders(path.join(w.dataDir, "runtimes"))).toEqual([original.id]);
    expect(fs.existsSync(baseline.directory)).toBe(false);
  });

  it("R11 a release that ships another CLI switches to it once everything is stopped, and the old CLI is removed", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", w);
    const changed = manifest(); changed.providers.copilot.version = "9.9.9";
    const install = async (prefix: string, spec: string) => { await w.install(prefix, spec); if (spec.startsWith("office-partner@")) w.put(path.join(prefix, "node_modules", "office-partner"), "bin/runtime-manifest.json", JSON.stringify(changed)); };
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, install, ...w.controls });
    const state = store.readState(w.dataDir);
    expect(state.profiles[state.active.copilot]).toMatchObject({ version: "9.9.9", stateHome: original.stateHome });
    expect(state.profiles[original.id]).toBeUndefined();
    await vi.waitFor(() => { expect(folders(path.join(w.dataDir, "runtimes"))).toEqual([state.active.copilot]); });
    expect(Object.keys(store.currentDeployment(w.dataDir).releases)).toEqual([updated.id]);
    expect(fs.existsSync(baseline.directory)).toBe(false);
  });

  it("a provider whose new CLI cannot be installed does not stop the update: its current CLI stays and the reason is recorded", async () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    const original = await installRuntime(w.root, "copilot", w);
    const changed = manifest(); changed.providers.copilot.version = "9.9.9";
    const install = async (prefix: string, spec: string) => {
      if (spec.startsWith("@github/copilot@9.9.9")) throw new Error("offline");
      await w.install(prefix, spec);
      if (spec.startsWith("office-partner@")) w.put(path.join(prefix, "node_modules", "office-partner"), "bin/runtime-manifest.json", JSON.stringify(changed));
    };
    const notice = vi.fn();
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, install, notice, ...w.controls });
    expect(w.calls).toEqual(["drain", "stop:1.0.0", "start:1.1.0"]);
    expect(store.currentDeployment(w.dataDir).active).toBe(updated.id);
    const state = store.readState(w.dataDir);
    expect(state.active.copilot).toBe(original.id);
    expect(state.failures.copilot).toMatchObject({ code: "RUNTIME_INSTALL_FAILED", version: "9.9.9" });
    expect(notice).toHaveBeenCalledWith(expect.stringContaining("copilot: the new CLI was not installed"));
  });

  it("a provider that is not signed in does not stop the update, and switches to its new CLI like any other", async () => {
    const w = fixture(); ensureBaseline(w.root, w.dataDir);
    await installRuntime(w.root, "copilot", w);
    const changed = manifest(); changed.providers.copilot.version = "9.9.9";
    const install = async (prefix: string, spec: string) => { await w.install(prefix, spec); if (spec.startsWith("office-partner@")) w.put(path.join(prefix, "node_modules", "office-partner"), "bin/runtime-manifest.json", JSON.stringify(changed)); };
    await updateRelease(w.root, "office-partner@1.1.0", { ...w, install, preflight: async () => { throw new Error("NATIVE_PREFLIGHT_AUTH_REQUIRED"); }, ...w.controls });
    const state = store.readState(w.dataDir);
    expect(state.profiles[state.active.copilot].version).toBe("9.9.9");
    expect(state.failures).toEqual({});
  });

  it("R12 drops installs that no provider runs, and sweeps folders left behind earlier", async () => {
    const w = fixture(), baseline = ensureBaseline(w.root, w.dataDir);
    await installRuntime(w.root, "copilot", w);
    const stray = await installRuntime(w.shipping("9.9.9"), "copilot", { ...w, activate: false });
    expect(store.readState(w.dataDir).profiles[stray.id]).toBeDefined();
    const updated = await updateRelease(w.root, "office-partner@1.1.0", { ...w, ...w.controls });
    expect(store.readState(w.dataDir).profiles[stray.id]).toBeUndefined();
    await vi.waitFor(() => { expect(fs.existsSync(path.dirname(stray.command))).toBe(false); });
    const orphan = path.join(w.dataDir, "releases", "a".repeat(64)); fs.mkdirSync(orphan, { recursive: true });
    const staging = path.join(w.dataDir, "releases", ".install-in-progress"); fs.mkdirSync(staging, { recursive: true });
    const swept = retireOldReleases(w.dataDir);
    expect(swept.releases).toEqual([]);
    await swept.pending;
    expect(fs.existsSync(orphan)).toBe(false); expect(fs.existsSync(staging)).toBe(true);
    expect(store.currentDeployment(w.dataDir).active).toBe(updated.id);
    expect(fs.existsSync(baseline.directory)).toBe(false);
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
    await installRuntime(w.root, "copilot", w);
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
});
