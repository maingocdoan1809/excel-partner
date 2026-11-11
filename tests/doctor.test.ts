import { afterEach, describe, expect, it, vi } from "vitest";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import http from "node:http";
import selfsigned from "selfsigned";
import { spawnSync } from "node:child_process";
// @ts-expect-error Standalone diagnostics ships as JavaScript, without the TS build.
import { recordStartup } from "../bin/startup-log.mjs";
// @ts-expect-error Standalone diagnostics ships as JavaScript, without the TS build.
import { collectDoctor, resolveCommands } from "../bin/doctor-checks.mjs";
// @ts-expect-error Standalone diagnostics ships as JavaScript, without the TS build.
import { doctorMain, parseDoctorArgs } from "../bin/doctor.mjs";
// @ts-expect-error Standalone diagnostics ships as JavaScript, without the TS build.
import { runCommand, requestHealth, probeMcp } from "../bin/doctor-probes.mjs";

const dirs: string[] = [];
afterEach(() => { for (const dir of dirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true }); });
function world(name = "office-partner") {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "partner-doctor-")); dirs.push(root);
  const put = (file: string, content: string) => {
    const target = path.join(root, file); fs.mkdirSync(path.dirname(target), { recursive: true }); fs.writeFileSync(target, content); return target;
  };
  put("package.json", JSON.stringify({ name, version: "0.1.22" }));
  put("node_modules/@github/copilot-sdk/package.json", JSON.stringify({ version: "1.0.18", copilotCliVersion: "1.0.94", engines: { node: "^20.19.0 || >=22.12.0" } }));
  put("dist/providers/src/copilot.js", 'static MIN_VERSION = "1.0.95";');
  put("dist/bridge/src/index.js", ""); put("gateway/dist/server.mjs", "");
  const exe = put(`private account/copilot${process.platform === "win32" ? ".exe" : ""}`, "mock executable"); fs.chmodSync(exe, 0o700);
  const env = { PATH: "", LOCALAPPDATA: root, AGENT_GATEWAY_HOME: path.join(root, "gateway-data"), COPILOT_BIN: exe, GITHUB_TOKEN: "sensitive-fixture-token" };
  const run = vi.fn(async () => ({ code: 0, stdout: "1.0.95\nprivate@example.invalid sensitive-fixture-token", stderr: "" }));
  const options = { root, env, target: "gateway", provider: "copilot", platform: "linux", nodeVersion: "22.12.0", run };
  return { root, put, exe, env, run, options };
}

describe("portable provider doctor", () => {
  describe("sign-in startup", () => {
    const registry = (node: string, launcher: string) => `\r\nHKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\r\n    Agent Gateway    REG_SZ    "${node}" "${launcher}" start gateway --at-logon\r\n`;
    const startupWorld = (launcherExists = true) => {
      const w = world();
      const node = w.put("tools/node.exe", ""), launcher = launcherExists ? w.put("Excel Partner/recovery/recovery.mjs", "") : path.join(w.root, "Excel Partner", "recovery", "recovery.mjs");
      const run = vi.fn(async (command: string) => command === "reg.exe" ? { code: 0, stdout: registry(node, launcher), stderr: "" } : { code: 0, stdout: "1.0.95", stderr: "" });
      return { ...w, dataDir: path.join(w.root, "Excel Partner"), options: { ...w.options, platform: "win32", run } };
    };
    const find = (report: { checks: Array<{ id: string }> }, id: string) => report.checks.find(check => check.id === id) as Record<string, any> | undefined;

    it("D5 passes when the entry still points to files that exist", async () => {
      const w = startupWorld();
      expect(find(await collectDoctor(w.options), "gateway.startup_target")).toMatchObject({ status: "pass", code: "STARTUP_TARGET_OK" });
    });

    it("D5 fails when Node.js or the launcher named by the entry has been moved or removed", async () => {
      const w = startupWorld(false);
      expect(find(await collectDoctor(w.options), "gateway.startup_target")).toMatchObject({ status: "fail", code: "STARTUP_TARGET_MISSING" });
    });

    it("D3 reports the last start at sign-in, with a reason code and no text from the machine", async () => {
      const w = startupWorld();
      expect(find(await collectDoctor(w.options), "gateway.startup_last")).toBeUndefined();
      recordStartup(w.dataDir, { service: "gateway", attempt: 3, ok: false, code: "NOT_ANSWERING", message: `Recorded gateway at ${w.root} is not answering` });
      const report = await collectDoctor(w.options);
      expect(find(report, "gateway.startup_last")).toMatchObject({ status: "warn", code: "STARTUP_LAST_FAILED", details: { attempt: 3, reason: "NOT_ANSWERING" }, action: expect.stringContaining("Fix connection") });
      expect(JSON.stringify(report)).not.toContain(w.root);
      recordStartup(w.dataDir, { service: "gateway", attempt: 1, ok: true });
      expect(find(await collectDoctor(w.options), "gateway.startup_last")).toMatchObject({ status: "pass", code: "STARTUP_LAST_OK" });
    });
  });

  it("H25 reports Node versions outside the pinned SQLite release's prebuilt range", async () => {
    const w = world(); w.put("package.json", JSON.stringify({ name: "office-partner", version: "0.1.22", engines: { node: "^20.19.0 || >=22.12.0 <26.0.0" } }));
    const report = await collectDoctor({ ...w.options, nodeVersion: "26.0.0" });
    expect(report.checks).toContainEqual(expect.objectContaining({ id: "installation.node", code: "NODE_UNSUPPORTED", status: "fail" }));
  });

  it("reports missing setup and unknown probes without starting providers", async () => {
    const w = world(); const report = await collectDoctor(w.options);
    expect(w.run).not.toHaveBeenCalled();
    expect(report).toMatchObject({ schemaVersion: 1, status: "fail", product: { name: "office-partner", version: "0.1.22" } });
    expect(report.checks).toContainEqual(expect.objectContaining({ id: "gateway.tls", code: "TLS_MISSING" }));
    expect(report.checks).toContainEqual(expect.objectContaining({ id: "provider.copilot.probe", code: "CLI_NOT_PROBED" }));
  });

  it("shows both SDK/runtime versions, suppresses updates only in the child environment, and exports no raw output", async () => {
    const w = world(); const report = await collectDoctor({ ...w.options, probe: true });
    expect(w.run).toHaveBeenCalledTimes(1);
    expect(w.run).toHaveBeenCalledWith(w.exe, ["--version"], expect.objectContaining({ env: expect.objectContaining({ COPILOT_AUTO_UPDATE: "false", DISABLE_UPDATES: "1" }) }));
    expect(w.env).not.toHaveProperty("COPILOT_AUTO_UPDATE");
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "SDK_CLI_PAIR_UNVERIFIED", details: { cliVersion: "1.0.95", sdkVersion: "1.0.18", bundledRuntime: "1.0.94" } }));
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "AUTH_UNVERIFIED" }));
    for (const secret of [w.root, "private account", "sensitive-fixture-token", "private@example.invalid"]) expect(JSON.stringify(report)).not.toContain(secret);
  });

  it.each([["0.0.1", "CLI_TOO_OLD"], ["nonsense sensitive-fixture-token", "CLI_VERSION_UNKNOWN"]])("classifies version %s", async (output, code) => {
    const w = world(); w.run.mockResolvedValue({ code: 0, stdout: output, stderr: "" });
    const report = await collectDoctor({ ...w.options, probe: true });
    expect(report.checks).toContainEqual(expect.objectContaining({ id: "provider.copilot.version", status: "fail", code }));
  });

  it("classifies CLI timeouts and missing executables without leaking exception text", async () => {
    const w = world();
    const report = await collectDoctor({ ...w.options, probe: true, run: async () => ({ code: null, errorCode: "ETIMEDOUT", stderr: w.exe + " sensitive-fixture-token" }) });
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "CLI_TIMEOUT" }));
    expect(JSON.stringify(report)).not.toContain("sensitive-fixture-token");
    const missing = await collectDoctor({ ...w.options, env: { ...w.env, COPILOT_BIN: path.join(w.root, "absent") } });
    expect(missing.checks).toContainEqual(expect.objectContaining({ code: "CLI_NOT_FOUND", status: "fail" }));
  });

  it("checks the SDK's tighter Node requirement independently from the product minimum", async () => {
    const w = world(); const report = await collectDoctor({ ...w.options, nodeVersion: "20.1.0" });
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "NODE_OK" }));
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "SDK_NODE_UNSUPPORTED", status: "fail" }));
  });

  it("does not interpret 'not logged in' as a successful Codex sign-in", async () => {
    const w = world();
    const run = vi.fn().mockResolvedValueOnce({ code: 0, stdout: "codex-cli 0.162.1" }).mockResolvedValueOnce({ code: 1, stderr: "Not logged in" });
    const report = await collectDoctor({ ...w.options, provider: "codex", env: { ...w.env, CODEX_BIN: w.exe }, probe: true, run });
    expect(report.checks).toContainEqual(expect.objectContaining({ id: "provider.codex.auth", status: "fail", code: "AUTH_REQUIRED" }));
  });

  it("checks service identity rather than treating every HTTP 200 as healthy", async () => {
    const w = world();
    const cert = await selfsigned.generate([{ name: "commonName", value: "localhost" }], { keySize: 2048 });
    w.put("gateway-data/localhost.crt", cert.cert); w.put("gateway-data/localhost.key", "key-presence-only");
    const health = vi.fn(async () => ({ gateway: false, bridge: true, ready: true }));
    const wrong = await collectDoctor({ ...w.options, health });
    expect(wrong.checks).toContainEqual(expect.objectContaining({ code: "SERVICE_IDENTITY_MISMATCH", status: "fail" }));
    const stopped = await collectDoctor({ ...w.options, health: async () => ({ errorCode: "ECONNREFUSED", raw: "sensitive-fixture-token" }) });
    expect(stopped.checks).toContainEqual(expect.objectContaining({ code: "SERVICE_STOPPED" }));
    expect(JSON.stringify(stopped)).not.toContain("sensitive-fixture-token");
    const ready = await collectDoctor({ ...w.options, health: async () => ({ gateway: true, ready: true }) });
    expect(ready.status).toBe("pass");
  });

  it("D01 identifies the responding gateway and distinguishes initialization from a stopped service", async () => {
    const w = world(); const cert = await selfsigned.generate([{ name: "commonName", value: "localhost" }], { keySize: 2048 });
    w.put("gateway-data/localhost.crt", cert.cert); w.put("gateway-data/localhost.key", "key-presence-only");
    const report = await collectDoctor({ ...w.options, health: async () => ({ gateway: true, ready: false }) });
    const check = report.checks.find((c: { id: string }) => c.id === "gateway.health");
    expect(check).toMatchObject({ status: "fail", code: "SERVICE_NOT_READY", details: { port: 8766 } });
    expect(check.message).toContain("Gateway on port 8766"); expect(check.message).toContain("initializing");
    expect(check.message).not.toContain("stopped");
  });

  it("detects protocol mismatch and invalid add-in configuration without exposing contents", async () => {
    const w = world();
    const cert = await selfsigned.generate([{ name: "commonName", value: "localhost" }], { keySize: 2048 });
    w.put("Excel Partner/localhost.crt", cert.cert); w.put("Excel Partner/localhost.key", "key-presence-only");
    w.put("Excel Partner/config.json", "invalid sensitive-fixture-token");
    w.put("dist/protocol/src/version.js", "export const PROTOCOL_VERSION = 1;");
    const report = await collectDoctor({ ...w.options, target: "addin", health: async () => ({ bridge: true, protocolVersion: 99 }) });
    expect(report.checks).toContainEqual(expect.objectContaining({ id: "addin.config", code: "CONFIG_INVALID" }));
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "PROTOCOL_MISMATCH", details: { expected: 1, actual: 99 } }));
    expect(JSON.stringify(report)).not.toContain("sensitive-fixture-token");
  });

  it("rejects damaged config but never exports arbitrary configuration fields", async () => {
    const w = world(); w.put("gateway-data/config.json", JSON.stringify({ enabledProviders: ["unknown-account"], secret: "sensitive-fixture-token" }));
    const report = await collectDoctor(w.options);
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "CONFIG_INVALID" }));
    expect(JSON.stringify(report)).not.toContain("sensitive-fixture-token");
  });

  it("makes MCP verification explicit and reports missing dependencies without claiming execution worked", async () => {
    const w = world();
    const checkMcp = vi.fn(async () => ({ count: 20, catalogHash: "a".repeat(64) }));
    await collectDoctor({ ...w.options, checkMcp }); expect(checkMcp).not.toHaveBeenCalled();
    const report = await collectDoctor({ ...w.options, mcp: true, host: "word", checkMcp });
    expect(report.checks).toContainEqual(expect.objectContaining({ code: "MCP_CATALOG_OK", details: { host: "word", toolCount: 20, catalogHash: "a".repeat(64) } }));
    const failed = await collectDoctor({ ...w.options, mcp: true, checkMcp: async () => { throw Object.assign(new Error(w.exe), { code: "ERR_MODULE_NOT_FOUND" }); } });
    expect(failed.checks).toContainEqual(expect.objectContaining({ code: "MCP_DEPENDENCY_MISSING" }));
    expect(JSON.stringify(failed)).not.toContain(w.root);
  });

  it("works with main's package name and preserves user reports when the output already exists", async () => {
    const w = world("excel-partner"); const write = vi.fn();
    const reportFile = path.join(w.root, "report.json");
    expect(await doctorMain(["gateway", "--json", "--output", reportFile], { ...w.options, write })).toBe(1);
    expect(JSON.parse(write.mock.calls[0]![0])).toMatchObject({ product: { name: "excel-partner" } });
    const saved = fs.readFileSync(reportFile, "utf8");
    await expect(doctorMain(["--output", reportFile], { ...w.options, write })).rejects.toThrow("Choose a new file");
    expect(fs.readFileSync(reportFile, "utf8")).toBe(saved);
  });

  it.each([["--provider"], ["--provider", "unknown"], ["--host", "outlook"], ["--probe", "--unexpected"], ["banana"]])("validates arguments %j", (...args) => {
    expect(() => parseDoctorArgs(args)).toThrow();
  });

  it("resolves an explicit executable with spaces and ignores empty PATH entries", () => {
    const w = world();
    expect(resolveCommands(w.exe, w.env)).toEqual([fs.realpathSync(w.exe)]);
    expect(resolveCommands("copilot", { PATH: "" })).toEqual([]);
  });

  it("runs the standalone doctor with no dependencies or compiled product", () => {
    const w = world();
    fs.rmSync(path.join(w.root, "node_modules"), { recursive: true });
    fs.rmSync(path.join(w.root, "dist"), { recursive: true });
    for (const name of ["doctor.mjs", "doctor-checks.mjs", "doctor-probes.mjs"]) w.put(`bin/${name}`, fs.readFileSync(path.join("bin", name), "utf8"));
    const result = spawnSync(process.execPath, [path.join(w.root, "bin", "doctor.mjs"), "gateway", "--json"], {
      env: { ...process.env, ...w.env, PATH: "", Path: "" }, encoding: "utf8", timeout: 15000, windowsHide: true,
    });
    expect(result.status, result.stderr).toBe(1);
    expect(JSON.parse(result.stdout)).toMatchObject({ schemaVersion: 1, status: "fail" });
  });
});

describe("bounded diagnostic transports", () => {
  it("initializes the compiled local MCP and lists schemas without connecting a document or invoking tools", async () => {
    const result = await probeMcp({ root: path.resolve("."), host: "excel" });
    expect(result.count).toBeGreaterThan(0);
    expect(result.catalogHash).toMatch(/^[a-f0-9]{64}$/);
  });
  it("bounds child output and time without returning raw failure output", async () => {
    const large = await runCommand(process.execPath, ["-e", "process.stdout.write('x'.repeat(70000))"]);
    expect(large).toMatchObject({ code: null, errorCode: "OUTPUT_LIMIT", stdout: "", stderr: "" });
    const timeout = await runCommand(process.execPath, ["-e", "setInterval(()=>{},1000)"], { timeoutMs: 100 });
    expect(timeout).toMatchObject({ code: null, errorCode: "ETIMEDOUT" });
  });

  it("extracts only service identity from local health, excluding document and account metadata", async () => {
    const server = http.createServer((_req, res) => res.end(JSON.stringify({ name: "excel-ai-addin-bridge", protocolVersion: 1, sessions: [{ title: "private workbook", token: "secret" }] })));
    await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
    try {
      const result = await requestHealth({ port: (server.address() as { port: number }).port });
      expect(result).toEqual({ bridge: true, gateway: false, ready: false, protocolVersion: 1 });
      expect(JSON.stringify(result)).not.toMatch(/private|secret/);
    } finally { await new Promise<void>(resolve => server.close(() => resolve())); }
  });
});
