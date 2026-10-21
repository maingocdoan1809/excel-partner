import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { RunPythonInputSchema, type RunPythonInput } from "../../protocol/src/index.js";
import { runLocalScript } from "./skill-runner.js";

export const PYTHON_FILE_BYTES = 10 * 1024 * 1024;
export class PythonError extends Error {}

/** Runs with the local user's permissions; the temporary folder is a workspace, not an isolation boundary. */
export async function runPython(
  input: RunPythonInput,
  files: Array<{ path: string; bytes: Buffer }>,
  signal?: AbortSignal,
) {
  input = RunPythonInputSchema.parse(input);
  if (signal?.aborted) throw new PythonError("Python execution cancelled.");
  if (files.reduce((sum, file) => sum + file.bytes.length, 0) > PYTHON_FILE_BYTES) throw new PythonError("Python inputs exceed 10 MB combined.");
  const root = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "excel-partner-python-")));
  const local = (name: string) => {
    const candidate = path.resolve(root, name);
    if (!candidate.startsWith(root + path.sep) || name.includes("\\") || name.split("/").some(part => !part || part === "." || part === ".." || /[:\x00-\x1f]/.test(part))) {
      throw new PythonError("Use a relative Python workspace path without traversal or drive letters.");
    }
    return candidate;
  };
  try {
    const seen = new Set<string>();
    for (const file of files) {
      const target = local(file.path);
      const key = process.platform === "win32" ? target.toLowerCase() : target;
      if (key === (process.platform === "win32" ? path.join(root, "main.py").toLowerCase() : path.join(root, "main.py")) || seen.has(key)) throw new PythonError("Duplicate or reserved Python input file path.");
      seen.add(key);
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.writeFileSync(target, file.bytes);
    }
    const file = path.join(root, "main.py");
    fs.writeFileSync(file, input.code, "utf8");
    const result = await runLocalScript({ root, file, python: true,
      interpreter: process.env.BRIDGE_PYTHON_BIN || (process.platform === "win32" ? "python" : "python3") }, input, signal);
    const outputFiles: Array<{ name: string; base64: string }> = [];
    if (result.exitCode === 0 && !result.timedOut && !result.cancelled && !result.outputTruncated) {
      let total = 0;
      for (const name of input.outputFiles) {
        const target = fs.realpathSync(local(name));
        if (!target.startsWith(root + path.sep) || !fs.statSync(target).isFile()) throw new PythonError("Output must be a regular file inside the Python workspace.");
        total += fs.statSync(target).size;
        if (total > PYTHON_FILE_BYTES) throw new PythonError("Python outputs exceed 10 MB combined.");
        outputFiles.push({ name, base64: fs.readFileSync(target).toString("base64") });
      }
    }
    return { ...result, outputFiles };
  } catch (err) {
    throw new PythonError(err instanceof Error ? err.message : String(err));
  } finally {
    // root is the canonical directory returned by mkdtempSync, never an input path.
    fs.rmSync(root, { recursive: true, force: true });
  }
}
