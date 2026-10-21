import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { parseToolInput } from "../protocol/src/index.js";
import { runPython } from "../bridge/src/python-runner.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { createAttachment } from "../bridge/src/attachments.js";

const input = (code: string, extra = {}) => parseToolInput("run_python", { code, ...extra });
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllEnvs(); });

function setup(approved = true) {
  const sessions = new SessionManager({ ttlMs: 60_000 });
  const session = sessions.create("python-test", null);
  vi.spyOn(session, "isConnected", "get").mockReturnValue(true);
  const permission = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved } });
  const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000);
  const call = (args: unknown) => gateway.invoke({ sessionId: session.id, tool: "run_python", args });
  return { session, permission, call };
}

describe("local Python execution", () => {
  it("rejects traversal, command options and excessive timeouts", () => {
    for (const output of ["../secret", "/etc/passwd", "C:/secret", "nested\\secret", "a/./b", "a//b"]) expect(() => input("print(1)", { outputFiles: [output] })).toThrow();
    expect(() => input("print(1)", { image: "anything" })).toThrow();
    expect(() => input("print(1)", { timeoutMs: 60_001 })).toThrow();
  });

  it("explains missing local Python", async () => {
    vi.stubEnv("BRIDGE_PYTHON_BIN", path.join(os.tmpdir(), "missing-python-" + Date.now()));
    await expect(runPython(input("print(1)"), [])).rejects.toThrow(/Install the interpreter.*BRIDGE_PYTHON_BIN/);
  });

  it("runs on the host with UTF-8 files, literal args and stdin, and cleans its working folder", async () => {
    const sentinel = fs.mkdtempSync(path.join(os.tmpdir(), "python-host-"));
    const hostFile = path.join(sentinel, "host.txt");
    fs.writeFileSync(hostFile, "local access");
    vi.stubEnv("EXCEL_SESSION_TOKEN", "private-credential");
    try {
      const result = await runPython(input([
        "import pathlib, sys, json, os, subprocess",
        `assert pathlib.Path(${JSON.stringify(hostFile)}).read_text() == 'local access'`,
        "assert 'EXCEL_SESSION_TOKEN' not in os.environ",
        "text = pathlib.Path('data/input.txt').read_text(encoding='utf-8')",
        "child = subprocess.run([sys.executable, '-c', 'print(' + repr(text) + ')'], capture_output=True, text=True, check=True)",
        "pathlib.Path('result.txt').write_text(child.stdout, encoding='utf-8')",
        "print(json.dumps([os.getcwd(), sys.argv[1:], sys.stdin.read(), text]))",
      ].join("\n"), { args: ["a b", "$(literal)", "Tiếng Việt"], stdin: "hello", outputFiles: ["result.txt"] }),
      [{ path: "data/input.txt", bytes: Buffer.from("日本語 Tiếng Việt") }]);
      expect(result).toMatchObject({ exitCode: 0, timedOut: false, cancelled: false, outputTruncated: false });
      const [cwd, args, stdin, text] = JSON.parse(result.stdout);
      expect([args, stdin, text]).toEqual([["a b", "$(literal)", "Tiếng Việt"], "hello", "日本語 Tiếng Việt"]);
      expect(Buffer.from(result.outputFiles[0]!.base64, "base64").toString().replace(/\r\n/g, "\n")).toBe("日本語 Tiếng Việt\n");
      expect(fs.existsSync(cwd)).toBe(false);
    } finally { fs.rmSync(sentinel, { recursive: true, force: true }); }
  });

  it("limits runtime and console output and returns script failures", async () => {
    const timeout = await runPython(input("import time; time.sleep(10)", { timeoutMs: 100 }), []);
    expect(timeout).toMatchObject({ timedOut: true, outputFiles: [] });
    const noisy = await runPython(input("while True: print('x' * 10000)"), []);
    expect(noisy.outputTruncated).toBe(true);
    expect(Buffer.byteLength(noisy.stdout + noisy.stderr)).toBeLessThanOrEqual(512 * 1024);
    const failed = await runPython(input("raise RuntimeError('expected failure')"), []);
    expect(failed.exitCode).not.toBe(0);
    expect(failed.stderr).toContain("expected failure");
  });

  it("rejects missing, oversized, and escaping output files", async () => {
    await expect(runPython(input("pass", { outputFiles: ["missing.txt"] }), [])).rejects.toThrow();
    await expect(runPython(input("open('big.txt', 'wb').write(b'x' * (10 * 1024 * 1024 + 1))", { outputFiles: ["big.txt"] }), [])).rejects.toThrow(/10 MB/);
    const outside = fs.mkdtempSync(path.join(os.tmpdir(), "python-output-"));
    fs.writeFileSync(path.join(outside, "private.txt"), "private");
    try {
      const code = process.platform === "win32"
        ? `import subprocess; subprocess.run(['cmd.exe', '/c', 'mklink', '/J', 'escape', ${JSON.stringify(outside)}], check=True)`
        : `import os; os.symlink(${JSON.stringify(outside)}, 'escape')`;
      await expect(runPython(input(code, { outputFiles: ["escape/private.txt"] }), [])).rejects.toThrow(/inside the Python workspace/);
    } finally { fs.rmSync(outside, { recursive: true, force: true }); }
  });

  it("does not run denied code and rejects unavailable inputs before approval", async () => {
    const { session, permission, call } = setup(false);
    expect(await call(input("raise Exception('must not run')"))).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(permission).toHaveBeenCalledWith("__authorize_python", input("raise Exception('must not run')"), 45_000);
    permission.mockClear();
    expect(await call(input("print(1)", { files: [{ path: "input.txt", attachment: "attachment://00000000-0000-0000-0000-000000000000" }] }))).toMatchObject({ ok: false, error: { message: expect.stringContaining("unavailable") } });
    expect(permission).not.toHaveBeenCalled();
    expect(session.scriptControllers.size).toBe(0);
  });

  it("returns generated attachments and cancels the process when its session closes", async () => {
    const { session, call } = setup();
    const file = createAttachment("selected.txt", Buffer.from("selected input"));
    session.attachments.set(file.id, file);
    expect(await call(input("from pathlib import Path; Path('out.txt').write_text(Path('in.txt').read_text())", {
      files: [{ path: "in.txt", attachment: `attachment://${file.id}` }], outputFiles: ["out.txt"],
    }))).toMatchObject({ ok: true, data: { execution: "local", exitCode: 0, outputFiles: [{ name: "out.txt", path: expect.stringMatching(/^attachment:\/\//) }] } });
    const running = call(input("import time; time.sleep(30)"));
    const timer = setTimeout(() => session.abortTurn(), 200);
    try { expect(await running).toMatchObject({ ok: false, error: { code: "cancelled" } }); }
    finally { clearTimeout(timer); }
    expect(session.scriptControllers.size).toBe(0);
  });
});
