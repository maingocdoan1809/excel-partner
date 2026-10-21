import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SkillStore } from "../bridge/src/skills.js";
import { resolveSkillScript, runSkillScript, windowsBash } from "../bridge/src/skill-runner.js";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { parseToolInput, SERVER_TOOLS } from "../protocol/src/index.js";

let root: string;
let store: SkillStore;
const write = (file: string, content: string) => {
  const target = path.join(root, "report", file);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, content);
};
const input = (file: string, extra = {}) => parseToolInput("run_skill_script", { name: "report", file, ...extra });

beforeEach(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), "skill-runner-"));
  write("SKILL.md", "---\nname: report\n---\nUse the scripts.");
  store = new SkillStore([root], path.join(root, "templates"));
});
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllEnvs(); fs.rmSync(root, { recursive: true, force: true }); });

describe("skill script execution", () => {
  it("finds per-user Git Bash before the Windows WSL bash launcher", () => {
    const gitBash = path.join(root, "Programs", "Git", "bin", "bash.exe");
    fs.mkdirSync(path.dirname(gitBash), { recursive: true });
    fs.writeFileSync(gitBash, "test executable");
    vi.stubEnv("LOCALAPPDATA", root);
    expect(windowsBash()).toBe(gitBash);
  });
  it("runs Python with literal arguments, stdin, skill cwd and no bridge credentials", async () => {
    vi.stubEnv("EXCEL_SESSION_TOKEN", "must-not-be-inherited");
    write("tools/run.py", 'import sys, json, os\nprint(json.dumps([sys.argv[1:], sys.stdin.read(), os.getcwd(), os.environ.get("EXCEL_SESSION_TOKEN")]))');
    const result = await runSkillScript(store, input("tools/run.py", { args: ["$(touch injected)", "a b", "Tiếng Việt"], stdin: '{"test":1}' }));
    expect(result).toMatchObject({ exitCode: 0, timedOut: false, cancelled: false, outputTruncated: false });
    expect(JSON.parse(result.stdout)).toEqual([["$(touch injected)", "a b", "Tiếng Việt"], '{"test":1}', path.join(root, "report"), null]);
    expect(fs.existsSync(path.join(root, "report", "injected"))).toBe(false);
  });

  it("runs Shell scripts without requiring executable permissions", async () => {
    write("tools/run.sh", 'printf "%s\\n" "$1"\nprintf "shell error\\n" >&2\nexit 7\n');
    const result = await runSkillScript(store, input("tools/run.sh", { args: ["a b; echo not executed"] }));
    expect(result).toMatchObject({ exitCode: 7, stdout: "a b; echo not executed\n", stderr: "shell error\n" });
  });

  it("uses UTF-8 for files and nested Python subprocess output", async () => {
    write("tools/unicode.py", [
      "import json, subprocess, sys, locale",
      "text = '日本語 Tiếng Việt'",
      "open('unicode.txt', 'w').write(text)",
      "child = subprocess.run([sys.executable, '-c', 'print(' + repr(text) + ')'], capture_output=True, text=True, check=True)",
      "print(json.dumps([sys.flags.utf8_mode, locale.getpreferredencoding(False), child.stdout.strip(), open('unicode.txt').read()], ensure_ascii=False))",
    ].join("\n"));
    const result = await runSkillScript(store, input("tools/unicode.py"));
    expect(result.exitCode).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual([1, "utf-8", "日本語 Tiếng Việt", "日本語 Tiếng Việt"]);
  });

  it("rejects traversal, absolute paths, unsupported scripts and escaping symlinks", () => {
    write("tools/run.py", "print('ok')");
    write("run.exe", "binary");
    fs.writeFileSync(path.join(root, "outside.py"), "print('outside')");
    fs.symlinkSync(path.join(root, "outside.py"), path.join(root, "report", "tools", "link.py"));
    for (const file of ["../outside.py", "/etc/passwd", "C:/run.py", "tools\\run.py", "run.exe", "tools/link.py", "tools/missing.py"]) {
      expect(() => resolveSkillScript(store, input(file))).toThrow();
    }
  });

  it("kills timed out scripts and bounds output", async () => {
    write("sleep.py", "import time\ntime.sleep(10)");
    const slow = await runSkillScript(store, input("sleep.py", { timeoutMs: 100 }));
    expect(slow.timedOut).toBe(true);
    expect(slow.exitCode).not.toBe(0);
    write("noisy.py", "while True: print('x' * 10000)");
    const noisy = await runSkillScript(store, input("noisy.py"));
    expect(noisy.outputTruncated).toBe(true);
    expect(Buffer.byteLength(noisy.stdout + noisy.stderr)).toBeLessThanOrEqual(512 * 1024);
  });

  it("stops a running script when the turn is cancelled", async () => {
    write("sleep.py", "import time\ntime.sleep(10)");
    const controller = new AbortController();
    const result = runSkillScript(store, input("sleep.py"), controller.signal);
    controller.abort();
    expect(await result).toMatchObject({ cancelled: true });
  });

  it("explains a missing interpreter and rejects invalid tool arguments", async () => {
    write("run.py", "print('ok')");
    vi.stubEnv("BRIDGE_PYTHON_BIN", path.join(root, "missing-python"));
    await expect(runSkillScript(store, input("run.py"))).rejects.toThrow(/BRIDGE_PYTHON_BIN/);
    expect(() => input("run.py", { timeoutMs: 60_001 })).toThrow();
    expect(() => input("run.py", { args: ["bad\0arg"] })).toThrow();
    expect(() => input("run.py", { command: "bash -c anything" })).toThrow();
  });
});

describe("gateway script permission and results", () => {
  function setup(approved: boolean) {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("script-test", null);
    vi.spyOn(session, "isConnected", "get").mockReturnValue(true);
    const permission = vi.spyOn(session, "requestTool").mockResolvedValue({ ok: true, data: { approved } });
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000, undefined, false, store);
    return { session, gateway, permission };
  }

  it("does not run a declined script", async () => {
    write("run.py", 'open("marker", "w").write("ran")');
    const { session, gateway, permission } = setup(false);
    const result = await gateway.invoke({ sessionId: session.id, tool: "run_skill_script", args: input("run.py") });
    expect(result).toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(permission).toHaveBeenCalledWith("__authorize_skill_script", input("run.py"), 45_000);
    expect(fs.existsSync(path.join(root, "report", "marker"))).toBe(false);
    expect(SERVER_TOOLS).toContain("run_skill_script");
  });

  it("runs after approval and returns generated files as readable attachments", async () => {
    write("run.py", 'open("result.json", "w").write(\'{"ok":true}\')\nprint("done")');
    const { session, gateway } = setup(true);
    const result = await gateway.invoke({ sessionId: session.id, tool: "run_skill_script", args: input("run.py", { outputFiles: ["result.json"] }) });
    expect(result).toMatchObject({ ok: true, data: { exitCode: 0, stdout: "done\n", outputFiles: [{ name: "result.json", path: expect.stringMatching(/^attachment:\/\//) }] } });
    const attachment = [...session.attachments.values()][0]!;
    const read = await gateway.invoke({ sessionId: session.id, tool: "read_attached_file", args: { path: `attachment://${attachment.id}` } });
    expect(read.ok).toBe(true);
    expect(attachment.bytes.toString()).toBe('{"ok":true}');
  });

  it("returns stderr and exit status as a failed tool result", async () => {
    write("run.py", 'import sys\nprint("missing dependency", file=sys.stderr)\nsys.exit(3)');
    const { session, gateway } = setup(true);
    expect(await gateway.invoke({ sessionId: session.id, tool: "run_skill_script", args: input("run.py") }))
      .toMatchObject({ ok: false, error: { detail: { exitCode: 3, stderr: "missing dependency\n" } } });
  });

  it("cancels an approved script when its session closes even without an active provider turn", async () => {
    write("run.py", "import time\ntime.sleep(10)");
    const { session, gateway } = setup(true);
    const running = gateway.invoke({ sessionId: session.id, tool: "run_skill_script", args: input("run.py") });
    const timer = setTimeout(() => session.abortTurn(), 100);
    try {
      expect(await running).toMatchObject({ ok: false, error: { code: "cancelled", detail: { cancelled: true } } });
      expect(session.scriptControllers.size).toBe(0);
    } finally { clearTimeout(timer); }
  });

  it("does not prompt for an invalid path or run after cancellation during approval", async () => {
    write("run.py", 'open("marker", "w").write("ran")');
    const { session, gateway, permission } = setup(true);
    expect((await gateway.invoke({ sessionId: session.id, tool: "run_skill_script", args: input("../outside.py") })).ok).toBe(false);
    expect(permission).not.toHaveBeenCalled();
    const controller = session.beginTurn("turn");
    permission.mockImplementation(async () => { controller.abort(); return { ok: true, data: { approved: true } }; });
    expect(await gateway.invoke({ sessionId: session.id, tool: "run_skill_script", args: input("run.py") }))
      .toMatchObject({ ok: false, error: { code: "cancelled" } });
    expect(fs.existsSync(path.join(root, "report", "marker"))).toBe(false);
  });
});
