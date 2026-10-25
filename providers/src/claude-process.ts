import type { ChildProcess } from "node:child_process";
import { randomUUID } from "node:crypto";
import { createInterface } from "node:readline";
import spawn from "cross-spawn";

interface PendingControl {
  resolve(value: unknown): void;
  reject(error: Error): void;
}

/** Claude's stream-json control protocol (distinct from JSON-RPC).
 * One reader lives for the entire process; a result ends a turn, not stdin.
 */
export class ClaudeProcess {
  readonly child: ChildProcess;
  private readonly pending = new Map<string, PendingControl>();
  private failure: Error | null = null;
  private closing: Promise<void> | null = null;
  private readonly stderr: string[] = [];

  constructor(options: {
    command: string; args: string[]; cwd: string; env: NodeJS.ProcessEnv;
    onMessage(message: Record<string, any>): void;
    onExit(error: Error): void;
    onRuntimeClosed?: () => void;
    onStderr(line: string): void;
  }) {
    this.child = spawn(options.command, options.args, {
      cwd: options.cwd, env: options.env, stdio: ["pipe", "pipe", "pipe"], windowsHide: true,
    });
    const stdout = createInterface({ input: this.child.stdout!, crlfDelay: Infinity });
    const stderr = createInterface({ input: this.child.stderr! });
    let notified = false;
    const markClosed = () => {
      if (notified) return;
      notified = true;
      options.onRuntimeClosed?.();
    };
    const fail = (error: Error) => {
      if (this.failure) return;
      this.failure = error;
      markClosed();
      for (const request of this.pending.values()) request.reject(error);
      this.pending.clear();
      stdout.close();
      stderr.close();
      options.onExit(error);
    };
    this.child.once("error", fail);
    // Invalidate reuse immediately, while allowing the stdout reader to drain.
    this.child.once("exit", markClosed);
    // 'exit' may precede the final stdout data. 'close' is emitted only after
    // stdio has closed, so result/usage records reach the active turn first.
    this.child.once("close", (code, signal) => fail(new Error(
      `claude exited (code=${code}, signal=${signal})${this.stderr.length ? `: ${this.stderr.slice(-3).join(" | ")}` : ""}`,
    )));
    this.child.stdin!.on("error", fail);
    stdout.on("line", line => {
      let message: Record<string, any>;
      try { message = JSON.parse(line); } catch { return; }
      if (!message || typeof message !== "object") return;
      if (message.type === "control_response") {
        const response = message.response;
        const request = this.pending.get(response?.request_id);
        if (!request) return;
        this.pending.delete(response.request_id);
        if (response.subtype === "error") request.reject(new Error(String(response.error ?? "Claude control request failed")));
        else request.resolve(response.response);
      } else options.onMessage(message);
    });
    stderr.on("line", line => {
      this.stderr.push(line);
      if (this.stderr.length > 40) this.stderr.shift();
      options.onStderr(line);
    });
  }

  get alive(): boolean { return !this.failure && !this.closing && this.child.exitCode === null && this.child.signalCode === null; }

  write(message: unknown): void {
    if (!this.alive || this.child.stdin!.destroyed) throw this.failure ?? new Error("Claude process is not running");
    this.child.stdin!.write(JSON.stringify(message) + "\n");
  }

  request(subtype: string, signal?: AbortSignal, timeoutMs = 30_000): Promise<unknown> {
    if (signal?.aborted) return Promise.reject(signal.reason);
    const id = randomUUID();
    return new Promise((resolve, reject) => {
      const finish = (error?: Error, value?: unknown) => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", abort);
        this.pending.delete(id);
        if (error) reject(error); else resolve(value);
      };
      const abort = () => finish(new Error(`Claude ${subtype} was cancelled`));
      const timer = setTimeout(() => finish(new Error(`Claude ${subtype} timed out`)), timeoutMs);
      this.pending.set(id, { resolve: value => finish(undefined, value), reject: error => finish(error) });
      signal?.addEventListener("abort", abort, { once: true });
      try { this.write({ type: "control_request", request_id: id, request: { subtype } }); }
      catch (error) { finish(error as Error); }
    });
  }

  close(): Promise<void> {
    return this.closing ??= this.terminate();
  }

  private async terminate(): Promise<void> {
    if (this.child.exitCode !== null || this.child.signalCode !== null) return;
    // npm command shims on Windows may own the actual CLI as a child process.
    if (process.platform === "win32" && this.child.pid) {
      await new Promise<void>(resolve => {
        const killer = spawn("taskkill", ["/PID", String(this.child.pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
        killer.once("error", () => resolve());
        killer.once("close", () => resolve());
      });
      if (this.child.exitCode !== null || this.child.signalCode !== null) return;
    }
    await new Promise<void>(resolve => {
      const timer = setTimeout(() => { this.child.kill("SIGKILL"); resolve(); }, 1500);
      this.child.once("close", () => { clearTimeout(timer); resolve(); });
      this.child.stdin?.end();
      this.child.kill("SIGTERM");
    });
  }
}
