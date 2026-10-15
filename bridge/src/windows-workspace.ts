import path from "node:path";
import fs from "node:fs";
import { execFile, spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { randomUUID } from "node:crypto";
import { promisify } from "node:util";
import { PROJECT_ROOT } from "./config.js";

const run = promisify(execFile);
type Mode = "pick" | "workbook";
const powershell = process.platform === "win32" ? "powershell.exe" :
  fs.existsSync("/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe")
    ? "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe" : "powershell.exe";

async function scriptPath(): Promise<string> {
  const local = path.join(PROJECT_ROOT, "scripts", "workspace-windows.ps1");
  if (process.platform === "win32") return local;
  const { stdout } = await run("wslpath", ["-w", local], { timeout: 5000 });
  return stdout.trim();
}

class WorkspaceWorker {
  private readonly pending = new Map<string, { resolve(value: string | null): void; reject(error: Error): void; timer: NodeJS.Timeout }>();
  private buffer = "";
  private readyResolve!: () => void;
  private readyReject!: (error: Error) => void;
  readonly ready = new Promise<void>((resolve, reject) => { this.readyResolve = resolve; this.readyReject = reject; });

  constructor(private readonly process: ChildProcessWithoutNullStreams) {
    process.stdout.setEncoding("utf8");
    process.stdout.on("data", (chunk: string) => this.receive(chunk));
    process.stderr.on("data", () => undefined);
    process.on("error", (error) => this.fail(error));
    process.on("exit", () => this.fail(new Error("Windows workspace helper stopped")));
  }

  private receive(chunk: string): void {
    this.buffer += chunk;
    let end: number;
    while ((end = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, end).trim();
      this.buffer = this.buffer.slice(end + 1);
      let reply: { ready?: boolean; id?: string; path?: string | null; error?: string | null };
      try { reply = JSON.parse(line); }
      catch { continue; }
      if (reply.ready) { this.readyResolve(); continue; }
      if (!reply.id) continue;
      const waiting = this.pending.get(reply.id);
      if (!waiting) continue;
      clearTimeout(waiting.timer);
      this.pending.delete(reply.id);
      if (reply.error) waiting.reject(new Error(reply.error));
      else waiting.resolve(reply.path ?? null);
    }
  }

  private fail(error: Error): void {
    this.readyReject(error);
    for (const waiting of this.pending.values()) {
      clearTimeout(waiting.timer);
      waiting.reject(error);
    }
    this.pending.clear();
    workerPromise = null;
  }

  request(mode: Mode, workbookName = ""): Promise<string | null> {
    const id = randomUUID();
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error("Windows workspace helper timed out"));
      }, mode === "pick" ? 10 * 60_000 : 15_000);
      this.pending.set(id, { resolve, reject, timer });
      this.process.stdin.write(JSON.stringify({ id, mode, workbookName }) + "\n");
    });
  }

  close(): void {
    this.process.stdin.end();
    this.process.kill();
  }
}

let workerPromise: Promise<WorkspaceWorker> | null = null;
let activeWorker: WorkspaceWorker | null = null;

async function ensureWorker(): Promise<WorkspaceWorker> {
  workerPromise ??= (async () => {
    const script = await scriptPath();
    const child = spawn(powershell, ["-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", script, "-Mode", "server"], { windowsHide: true });
    const worker = new WorkspaceWorker(child);
    const timeout = setTimeout(() => child.kill(), 15_000);
    try { await worker.ready; activeWorker = worker; return worker; }
    finally { clearTimeout(timeout); }
  })();
  return workerPromise;
}

/** Start Windows and WinForms while the pane loads, so the folder dialog opens promptly on click. */
export function warmWorkspaceHelper(): void {
  void ensureWorker().catch(() => undefined);
}

export function stopWorkspaceHelper(): void {
  activeWorker?.close();
  activeWorker = null;
  workerPromise = null;
}

async function callOnce(mode: Mode, workbookName: string): Promise<string | null> {
  const script = await scriptPath();
  const { stdout } = await run(powershell, [
    "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", script,
    "-Mode", mode, "-WorkbookName", workbookName,
  ], { timeout: mode === "pick" ? 10 * 60_000 : 15_000, maxBuffer: 32 * 1024, windowsHide: true });
  return stdout.trim() || null;
}

async function callWindows(mode: Mode, workbookName = ""): Promise<string | null> {
  try { return await (await ensureWorker()).request(mode, workbookName); }
  catch { return callOnce(mode, workbookName); }
}

export async function pickWorkspaceFolder(): Promise<string | null> {
  return callWindows("pick");
}

export async function activeWorkbookPath(workbookName: string): Promise<string | null> {
  return callWindows("workbook", workbookName);
}
