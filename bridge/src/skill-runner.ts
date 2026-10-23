import fs from "node:fs";
import path from "node:path";
import spawn from "cross-spawn";
import type { RunSkillScriptInput } from "../../protocol/src/index.js";
import { SkillError, type SkillStore } from "./skills.js";

const MAX_OUTPUT_BYTES = 512 * 1024;

/** Windows' system bash.exe may launch WSL; prefer an installed native Git Bash. */
export function windowsBash(): string {
  const candidates = [
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, "Programs", "Git", "bin", "bash.exe"),
    process.env.ProgramFiles && path.join(process.env.ProgramFiles, "Git", "bin", "bash.exe"),
    process.env["ProgramFiles(x86)"] && path.join(process.env["ProgramFiles(x86)"]!, "Git", "bin", "bash.exe"),
  ];
  return candidates.find((candidate): candidate is string => Boolean(candidate && fs.existsSync(candidate))) ?? "bash";
}

export function skillFile(store: SkillStore, name: string, relative: string): { root: string; file: string } {
  store.requireEnabled(name);
  const skill = store.get(name);
  if (!skill) throw new SkillError(`unknown skill '${name}'`);
  if (!relative || relative.includes("\\") || relative.split("/").some(part => !part || part === "." || part === ".." || part.includes(":"))) {
    throw new SkillError("Choose a relative file inside the skill folder.");
  }
  try {
    const root = fs.realpathSync(skill.dir);
    const file = fs.realpathSync(path.resolve(root, relative));
    if (!file.startsWith(root + path.sep) || !fs.statSync(file).isFile()) throw new Error("outside");
    return { root, file };
  } catch {
    throw new SkillError(`'${relative}' is not a file inside skill '${name}'.`);
  }
}

export function resolveSkillScript(store: SkillStore, input: RunSkillScriptInput) {
  const selected = skillFile(store, input.name, input.file);
  const ext = path.extname(selected.file).toLowerCase();
  if (![".py", ".sh"].includes(ext) || path.extname(input.file).toLowerCase() !== ext) {
    throw new SkillError("Only Python (.py) and Bash (.sh) skill scripts can be run.");
  }
  return { ...selected, interpreter: ext === ".py"
    ? process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3")
    : process.env.BRIDGE_BASH_BIN || (process.platform === "win32" ? windowsBash() : "bash"), python: ext === ".py" };
}

export interface SkillScriptResult {
  name: string;
  file: string;
  exitCode: number | null;
  signal: string | null;
  stdout: string;
  stderr: string;
  timedOut: boolean;
  cancelled: boolean;
  outputTruncated: boolean;
}

/** Inherited OS permissions apply: the path check limits entry points, not what trusted scripts can do. */
export async function runSkillScript(store: SkillStore, input: RunSkillScriptInput, signal?: AbortSignal): Promise<SkillScriptResult> {
  const selected = resolveSkillScript(store, input);
  return { name: input.name, file: input.file, ...await runLocalScript(selected, input, signal) };
}

export async function runLocalScript(
  selected: { root: string; file: string; interpreter: string; python: boolean },
  input: Pick<RunSkillScriptInput, "args" | "stdin" | "timeoutMs">,
  signal?: AbortSignal,
): Promise<Omit<SkillScriptResult, "name" | "file">> {
  if (signal?.aborted) throw new SkillError("Script execution cancelled.");
  // Do not pass bridge/provider credentials to child processes.
  const env: NodeJS.ProcessEnv = {};
  for (const key of Object.keys(process.env)) {
    if (/^(path|pathext|systemroot|windir|comspec|home|userprofile|homedrive|homepath|temp|tmp|tmpdir|localappdata|appdata|lang|lc_.*|virtual_env)$/i.test(key)) env[key] = process.env[key];
  }
  env.PYTHONIOENCODING = "utf-8";
  // Also covers locale-based decoding in Python subprocess.run(text=True).
  env.PYTHONUTF8 = "1";
  env.PYTHONDONTWRITEBYTECODE = "1";
  return new Promise((resolve, reject) => {
    const child = spawn(selected.interpreter, selected.python
      ? ["-B", "-u", selected.file, ...input.args]
      : ["--noprofile", "--norc", selected.file.replace(/\\/g, "/"), ...input.args], {
      cwd: selected.root, env, shell: false, windowsHide: true,
      detached: process.platform !== "win32", stdio: ["pipe", "pipe", "pipe"],
    });
    let timedOut = false;
    let cancelled = false;
    let outputTruncated = false;
    let outputBytes = 0;
    const stdout: Buffer[] = [];
    const stderr: Buffer[] = [];
    let stopped = false;
    const stop = () => {
      if (stopped) return;
      stopped = true;
      if (!child.pid) return;
      if (process.platform === "win32") {
        const killer = spawn("taskkill", ["/pid", String(child.pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
        killer.once("error", () => child.kill("SIGKILL"));
        killer.once("close", code => { if (code !== 0) child.kill("SIGKILL"); });
      } else {
        try { process.kill(-child.pid, "SIGKILL"); } catch { child.kill("SIGKILL"); }
      }
    };
    const collect = (chunks: Buffer[], chunk: Buffer) => {
      const available = Math.max(0, MAX_OUTPUT_BYTES - outputBytes);
      if (available) chunks.push(chunk.subarray(0, available));
      outputBytes += chunk.length;
      if (outputBytes > MAX_OUTPUT_BYTES) { outputTruncated = true; stop(); }
    };
    child.stdout?.on("data", (chunk: Buffer) => collect(stdout, chunk));
    child.stderr?.on("data", (chunk: Buffer) => collect(stderr, chunk));
    const timer = setTimeout(() => { timedOut = true; stop(); }, input.timeoutMs);
    const abort = () => { cancelled = true; stop(); };
    signal?.addEventListener("abort", abort, { once: true });
    const cleanup = () => { clearTimeout(timer); signal?.removeEventListener("abort", abort); };
    child.once("error", error => {
      cleanup();
      reject(new SkillError(`Could not run ${selected.python ? "Python" : "Bash"}: ${error.message}. Install the interpreter on the bridge machine or set ${selected.python ? "BRIDGE_PYTHON_BIN" : "BRIDGE_BASH_BIN"} to its executable path.`));
    });
    child.once("close", (exitCode, exitSignal) => {
      cleanup();
      resolve({ exitCode, signal: exitSignal,
        stdout: Buffer.concat(stdout).toString("utf8"), stderr: Buffer.concat(stderr).toString("utf8"),
        timedOut, cancelled, outputTruncated });
    });
    // A script may exit without consuming stdin.
    child.stdin?.on("error", () => undefined);
    child.stdin?.end(input.stdin ?? "");
    if (signal?.aborted) abort();
  });
}
