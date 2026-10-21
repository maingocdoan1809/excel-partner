import { randomUUID } from "node:crypto";
import spawn from "cross-spawn";
import { z } from "zod";
import type { RunPythonInput } from "../../protocol/src/index.js";

export const SANDBOX_FILE_BYTES = 10 * 1024 * 1024;
const MAX_RESPONSE_BYTES = 16 * 1024 * 1024;
export class SandboxError extends Error {}

// The wrapper and user code both live only in the container. Nothing is mounted from the host.
const BOOTSTRAP = String.raw`
import base64, json, os, pathlib, stat, subprocess, sys
root = pathlib.Path('/work')
payload = json.load(sys.stdin)
def local(name):
    p = (root / name).resolve()
    if not p.is_relative_to(root) or p == root:
        raise ValueError('Invalid sandbox path')
    return p
try:
    for f in payload['files']:
        p = local(f['path'])
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(base64.b64decode(f['base64'], validate=True))
    (root / 'main.py').write_text(payload['code'], encoding='utf-8')
    timed_out = False
    try:
        r = subprocess.run([sys.executable, '-I', '-u', '/work/main.py', *payload['args']],
            input=payload['stdin'].encode('utf-8'), capture_output=True, timeout=payload['timeoutMs']/1000)
        out, err, code = r.stdout, r.stderr, r.returncode
    except subprocess.TimeoutExpired as e:
        out, err, code, timed_out = e.stdout or b'', e.stderr or b'', None, True
    limit = 512 * 1024
    truncated = len(out) + len(err) > limit
    out, err = out[:limit], err[:max(0, limit-len(out))]
    result = dict(exitCode=code, stdout=out.decode('utf-8', 'replace'), stderr=err.decode('utf-8', 'replace'),
        timedOut=timed_out, outputTruncated=truncated, outputFiles=[])
    if code == 0 and not timed_out and not truncated:
        total = 0
        for name in payload['outputFiles']:
            p = local(name)
            if not stat.S_ISREG(p.stat().st_mode):
                raise ValueError('Output must be a regular file: ' + name)
            total += p.stat().st_size
            if total > 10 * 1024 * 1024:
                raise ValueError('Generated files exceed 10 MB combined')
            result['outputFiles'].append(dict(name=name, base64=base64.b64encode(p.read_bytes()).decode('ascii')))
    print(json.dumps(result, ensure_ascii=True))
except Exception as e:
    print(json.dumps(dict(exitCode=1, stdout='', stderr=str(e), timedOut=False, outputTruncated=False, outputFiles=[])))
`;

const ResponseSchema = z.object({
  exitCode: z.number().int().nullable(), stdout: z.string(), stderr: z.string(),
  timedOut: z.boolean(), outputTruncated: z.boolean(),
  outputFiles: z.array(z.object({ name: z.string(), base64: z.string() })).max(8),
});
export type PythonSandboxResult = z.infer<typeof ResponseSchema>;

export function sandboxArgs(name: string, image: string): string[] {
  return ["run", "--rm", "--pull=never", "--name", name, "--label", "com.excel-partner.sandbox=true",
    "--network=none", "--read-only", "--cap-drop=ALL", "--security-opt=no-new-privileges",
    "--user=65534:65534", "--memory=256m", "--memory-swap=256m", "--cpus=1", "--pids-limit=64",
    "--ulimit", "nofile=128:128", "--ulimit", "fsize=16777216:16777216",
    "--tmpfs", "/work:rw,nosuid,nodev,noexec,size=32m,mode=1777",
    "--tmpfs", "/tmp:rw,nosuid,nodev,noexec,size=16m,mode=1777", "--workdir=/work",
    "--env=PYTHONUTF8=1", "--env=PYTHONIOENCODING=utf-8", "--env=PYTHONDONTWRITEBYTECODE=1",
    "--entrypoint=python", "-i", image, "-I", "-u", "-c", BOOTSTRAP];
}

function removeContainer(command: string, name: string): Promise<void> {
  return new Promise(resolve => {
    const child = spawn(command, ["rm", "--force", name], { windowsHide: true, shell: false, stdio: "ignore" });
    const timer = setTimeout(() => { child.kill(); resolve(); }, 3000);
    const done = () => { clearTimeout(timer); resolve(); };
    child.once("error", done);
    child.once("close", done);
  });
}

export async function runPythonSandbox(
  input: RunPythonInput,
  files: Array<{ path: string; bytes: Buffer }>,
  signal?: AbortSignal,
): Promise<PythonSandboxResult> {
  if (signal?.aborted) throw new SandboxError("Python execution cancelled.");
  const command = process.env.BRIDGE_DOCKER_BIN || "docker";
  const image = process.env.BRIDGE_PYTHON_SANDBOX_IMAGE || "python:3.13-slim";
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._/:@-]{0,255}$/.test(image)) throw new SandboxError("Invalid Python sandbox image.");
  if (files.reduce((sum, file) => sum + file.bytes.length, 0) > SANDBOX_FILE_BYTES) throw new SandboxError("Sandbox inputs exceed 10 MB combined.");
  const name = `excel-partner-python-${randomUUID()}`;
  let timedOut = false;
  let oversized = false;
  const abort = () => { void removeContainer(command, name); child.kill(); };
  const child = spawn(command, sandboxArgs(name, image), { shell: false, windowsHide: true, stdio: ["pipe", "pipe", "pipe"] });
  const timer = setTimeout(() => { timedOut = true; abort(); }, input.timeoutMs + 10_000);
  signal?.addEventListener("abort", abort, { once: true });
  if (signal?.aborted) abort();
  try {
    const response = await new Promise<{ code: number | null; stdout: string; stderr: string }>((resolve, reject) => {
      let size = 0;
      const stdout: Buffer[] = [];
      const stderr: Buffer[] = [];
      const collect = (chunks: Buffer[], chunk: Buffer) => {
        const remaining = Math.max(0, MAX_RESPONSE_BYTES - size);
        if (remaining) chunks.push(chunk.subarray(0, remaining));
        size += chunk.length;
        if (size > MAX_RESPONSE_BYTES && !oversized) { oversized = true; abort(); }
      };
      child.stdout?.on("data", (chunk: Buffer) => collect(stdout, chunk));
      child.stderr?.on("data", (chunk: Buffer) => collect(stderr, chunk));
      child.once("error", error => reject(new SandboxError(`Docker sandbox unavailable: ${error.message}. Install/start Docker and prepare the Python image. Local execution is disabled for this tool.`)));
      child.once("close", code => resolve({ code, stdout: Buffer.concat(stdout).toString("utf8"), stderr: Buffer.concat(stderr).toString("utf8") }));
      child.stdin?.on("error", () => undefined);
      child.stdin?.end(JSON.stringify({ ...input, files: files.map(file => ({ path: file.path, base64: file.bytes.toString("base64") })) }));
    });
    if (signal?.aborted) throw new SandboxError("Python execution cancelled.");
    if (timedOut) throw new SandboxError("Python sandbox timed out.");
    if (oversized) throw new SandboxError("Python sandbox exceeded the response size limit.");
    if (response.code !== 0) throw new SandboxError(`Docker sandbox failed (exit ${response.code}): ${response.stderr.slice(0, 2000) || "container stopped; memory or process limit may have been exceeded"}. No local execution was attempted.`);
    const result = ResponseSchema.safeParse(JSON.parse(response.stdout));
    if (!result.success) throw new SandboxError("Invalid response from Python sandbox.");
    let outputBytes = 0;
    for (const file of result.data.outputFiles) {
      if (!input.outputFiles.includes(file.name)) throw new SandboxError("Sandbox returned an unrequested output file.");
      outputBytes += Buffer.from(file.base64, "base64").length;
    }
    if (outputBytes > SANDBOX_FILE_BYTES) throw new SandboxError("Sandbox outputs exceed 10 MB combined.");
    return result.data;
  } catch (err) {
    if (err instanceof SandboxError) throw err;
    throw new SandboxError(`Python sandbox failed: ${err instanceof Error ? err.message : String(err)}`);
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener("abort", abort);
    await removeContainer(command, name);
  }
}
