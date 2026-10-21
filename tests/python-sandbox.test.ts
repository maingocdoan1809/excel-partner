import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { parseToolInput } from "../protocol/src/index.js";
import { runPythonSandbox, sandboxArgs } from "../bridge/src/python-sandbox.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { createAttachment } from "../bridge/src/attachments.js";

const input = (code: string, extra = {}) => parseToolInput("run_python", { code, ...extra });
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllEnvs(); });

describe("Python sandbox contract", () => {
  it("has no host mounts, network, root user or writable system and caps resources", () => {
    const args = sandboxArgs("test-container", "python:3.13-slim");
    for (const flag of ["--network=none", "--read-only", "--cap-drop=ALL", "--security-opt=no-new-privileges", "--user=65534:65534", "--memory=256m", "--memory-swap=256m", "--cpus=1", "--pids-limit=64", "--pull=never"]) expect(args).toContain(flag);
    expect(args.some(arg => /^--(?:volume|mount|privileged|pid=host|network=host)/.test(arg) || arg === "-v")).toBe(false);
  });

  it("rejects input/output paths outside the sandbox and arbitrary command options", () => {
    for (const output of ["../secret", "/etc/passwd", "C:/secret", "nested\\secret", "a/./b", "a//b"]) expect(() => input("print(1)", { outputFiles: [output] })).toThrow();
    expect(() => input("print(1)", { image: "untrusted-image" })).toThrow();
    expect(() => input("print(1)", { timeoutMs: 60_001 })).toThrow();
  });

  it("fails closed when Docker is missing", async () => {
    vi.stubEnv("BRIDGE_DOCKER_BIN", path.join(os.tmpdir(), "missing-docker-" + Date.now()));
    await expect(runPythonSandbox(input("raise Exception('must not run on host')"), [])).rejects.toThrow(/Docker sandbox unavailable/);
  });

  it("does not run denied code and rejects unavailable attachments before asking", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("python-test", null);
    vi.spyOn(session, "isConnected", "get").mockReturnValue(true);
    const permission = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: false } });
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000);
    expect(await gateway.invoke({ sessionId: session.id, tool: "run_python", args: input("print(1)") })).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(permission).toHaveBeenCalledWith("__authorize_python", input("print(1)"), 45_000);
    permission.mockClear();
    expect(await gateway.invoke({ sessionId: session.id, tool: "run_python", args: input("print(1)", { files: [{ path: "input.txt", attachment: "attachment://00000000-0000-0000-0000-000000000000" }] }) }))
      .toMatchObject({ ok: false, error: { message: expect.stringContaining("unavailable") } });
    expect(permission).not.toHaveBeenCalled();
    expect(session.scriptControllers.size).toBe(0);
  });
});

describe.skipIf(process.env.RUN_DOCKER_TESTS !== "1")("Python sandbox (real Docker)", () => {
  it("creates and runs Python files, copies only supplied inputs and returns UTF-8 output files", async () => {
    const code = [
      "import pathlib, subprocess, sys",
      "text = pathlib.Path('data/input.txt').read_text(encoding='utf-8')",
      "pathlib.Path('generated.py').write_text('print(' + repr(text) + ')', encoding='utf-8')",
      "r = subprocess.run([sys.executable, 'generated.py'], capture_output=True, text=True, check=True)",
      "print(r.stdout, end='')",
      "pathlib.Path('result.txt').write_text(r.stdout, encoding='utf-8')",
    ].join("\n");
    const result = await runPythonSandbox(input(code, { outputFiles: ["generated.py", "result.txt"] }), [{ path: "data/input.txt", bytes: Buffer.from("日本語 Tiếng Việt") }]);
    expect(result).toMatchObject({ exitCode: 0, stdout: "日本語 Tiếng Việt\n", timedOut: false, outputTruncated: false });
    expect(Buffer.from(result.outputFiles[1]!.base64, "base64").toString()).toBe("日本語 Tiếng Việt\n");
  });

  it("cannot write the system, access host files, inherit secrets or connect to a network", async () => {
    const sentinel = path.join(os.tmpdir(), `sandbox-host-only-${Date.now()}.txt`);
    fs.writeFileSync(sentinel, "host private data");
    vi.stubEnv("EXCEL_SESSION_TOKEN", "must-not-enter-container");
    try {
      const code = [
        "import os, pathlib, socket, json",
        `assert not pathlib.Path(${JSON.stringify(sentinel)}).exists()`,
        "assert os.getuid() == 65534",
        "assert 'EXCEL_SESSION_TOKEN' not in os.environ",
        "try: pathlib.Path('/etc/sandbox-test').write_text('no')",
        "except OSError: pass",
        "else: raise AssertionError('system writable')",
        "try: socket.create_connection(('1.1.1.1', 443), timeout=1)",
        "except OSError: pass",
        "else: raise AssertionError('network available')",
        "print('isolated')",
      ].join("\n");
      expect(await runPythonSandbox(input(code), [])).toMatchObject({ exitCode: 0, stdout: "isolated\n" });
    } finally { fs.unlinkSync(sentinel); }
  });

  it("times out code and refuses output files pointing outside /work", async () => {
    expect(await runPythonSandbox(input("import time; time.sleep(10)", { timeoutMs: 100 }), [])).toMatchObject({ timedOut: true, exitCode: null });
    const result = await runPythonSandbox(input("import os; os.symlink('/etc/passwd', 'leak.txt')", { outputFiles: ["leak.txt"] }), []);
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain("Invalid sandbox path");
    expect(result.outputFiles).toEqual([]);
  });

  it("gateway returns generated attachments and cancellation removes the running container", async () => {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("python-docker-test", null);
    vi.spyOn(session, "isConnected", "get").mockReturnValue(true);
    vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved: true } });
    const file = createAttachment("selected.txt", Buffer.from("selected input"));
    session.attachments.set(file.id, file);
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000);
    const result = await gateway.invoke({ sessionId: session.id, tool: "run_python", args: input("from pathlib import Path; Path('out.txt').write_text(Path('in.txt').read_text())", {
      files: [{ path: "in.txt", attachment: `attachment://${file.id}` }], outputFiles: ["out.txt"],
    }) });
    expect(result).toMatchObject({ ok: true, data: { sandbox: "docker", exitCode: 0, outputFiles: [{ name: "out.txt", path: expect.stringMatching(/^attachment:\/\//) }] } });
    const running = gateway.invoke({ sessionId: session.id, tool: "run_python", args: input("import time; time.sleep(30)") });
    const timer = setTimeout(() => session.abortTurn(), 1500);
    try { expect(await running).toMatchObject({ ok: false, error: { code: "cancelled" } }); }
    finally { clearTimeout(timer); }
    expect(session.scriptControllers.size).toBe(0);
  });
});
