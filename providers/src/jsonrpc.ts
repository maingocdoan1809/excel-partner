import type { ChildProcess } from "node:child_process";
import { createInterface } from "node:readline";
import spawn from "cross-spawn";

export interface JsonRpcOptions {
  command: string;
  args: string[];
  env?: NodeJS.ProcessEnv;
  cwd?: string;
  /** ACP requires `"jsonrpc":"2.0"`; Codex app-server omits it. */
  versionField?: boolean;
  /** Receives stderr lines from the child process. */
  onStderr?: (line: string) => void;
  onNotification?: (method: string, params: unknown) => void;
  /**
   * Handle a request sent by the child (approvals, permission prompts, ...).
   * Throw to answer with a JSON-RPC error.
   */
  onRequest?: (method: string, params: unknown, id: number | string) => Promise<unknown> | unknown;
}

export class JsonRpcError extends Error {
  constructor(
    readonly code: number,
    message: string,
    readonly data?: unknown,
  ) {
    super(message);
    this.name = "JsonRpcError";
  }
}

interface Pending {
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  method: string;
}

/**
 * Newline-delimited JSON-RPC peer over a child process' stdio. Shared by the
 * Codex app-server and Copilot ACP adapters.
 */
export class JsonRpcPeer {
  readonly child: ChildProcess;
  private nextId = 1;
  private readonly pending = new Map<number | string, Pending>();
  private closed = false;
  private readonly exitListeners: Array<(code: number | null, signal: NodeJS.Signals | null) => void> = [];

  exitInfo: { code: number | null; signal: NodeJS.Signals | null } | null = null;
  stderrTail: string[] = [];

  constructor(private readonly options: JsonRpcOptions) {
    this.child = spawn(options.command, options.args, {
      cwd: options.cwd,
      env: options.env ?? process.env,
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
    });

    createInterface({ input: this.child.stdout! }).on("line", (line) => this.onLine(line));
    createInterface({ input: this.child.stderr! }).on("line", (line) => {
      this.stderrTail.push(line);
      if (this.stderrTail.length > 40) this.stderrTail.shift();
      options.onStderr?.(line);
    });

    // A broken pipe on write must never crash the bridge.
    this.child.stdin!.on("error", () => undefined);

    this.child.once("error", (error) => {
      // Spawn failures (ENOENT, EACCES) never emit `exit`; treat them as one.
      if (!this.exitInfo) {
        this.exitInfo = { code: null, signal: null };
        this.teardown(error);
        for (const listener of this.exitListeners) listener(null, null);
      }
    });
    this.child.once("exit", (code, signal) => {
      this.exitInfo = { code, signal };
      this.teardown(new Error(`process exited (code=${code}, signal=${signal})`));
      for (const listener of this.exitListeners) listener(code, signal);
    });
  }

  get alive(): boolean {
    return !this.closed;
  }

  onExit(listener: (code: number | null, signal: NodeJS.Signals | null) => void): void {
    if (this.exitInfo) listener(this.exitInfo.code, this.exitInfo.signal);
    else this.exitListeners.push(listener);
  }

  request<T = unknown>(method: string, params?: unknown, timeoutMs = 60_000): Promise<T> {
    if (this.closed) return Promise.reject(new Error(`${method}: process is not running`));
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`${method}: timed out after ${timeoutMs}ms`));
      }, timeoutMs);
      this.pending.set(id, {
        method,
        resolve: (value) => {
          clearTimeout(timer);
          resolve(value as T);
        },
        reject: (error) => {
          clearTimeout(timer);
          reject(error);
        },
      });
      this.write({ id, method, params });
    });
  }

  notify(method: string, params?: unknown): void {
    if (this.closed) return;
    this.write(params === undefined ? { method } : { method, params });
  }

  /** Terminate the child; escalates to SIGKILL if it ignores SIGTERM. */
  async close(graceMs = 1500): Promise<void> {
    if (this.exitInfo) return;
    if (process.platform === "win32" && this.child.pid) {
      // npm .cmd launchers own the real CLI as a grandchild. Killing only cmd.exe leaks it.
      await new Promise<void>(resolve => {
        const killer = spawn("taskkill", ["/PID", String(this.child.pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
        killer.once("error", () => resolve()); killer.once("close", () => resolve());
      });
      if (this.exitInfo) return;
    }
    this.child.stdin?.end();
    this.child.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      const timer = setTimeout(() => {
        this.child.kill("SIGKILL");
        resolve();
      }, graceMs);
      this.onExit(() => {
        clearTimeout(timer);
        resolve();
      });
    });
  }

  private write(message: Record<string, unknown>): void {
    const payload = this.options.versionField ? { jsonrpc: "2.0", ...message } : message;
    this.child.stdin!.write(JSON.stringify(payload) + "\n");
  }

  private onLine(line: string): void {
    const trimmed = line.trim();
    if (!trimmed.startsWith("{")) return;
    let message: Record<string, unknown>;
    try {
      message = JSON.parse(trimmed) as Record<string, unknown>;
    } catch {
      return;
    }

    const id = message.id as number | string | undefined;
    const method = message.method as string | undefined;

    if (method !== undefined && id !== undefined) {
      void this.answerRequest(id, method, message.params);
      return;
    }
    if (method !== undefined) {
      this.options.onNotification?.(method, message.params);
      return;
    }
    if (id !== undefined) {
      const pending = this.pending.get(id);
      if (!pending) return;
      this.pending.delete(id);
      const error = message.error as { code?: number; message?: string; data?: unknown } | undefined;
      if (error) {
        pending.reject(new JsonRpcError(error.code ?? -32000, error.message ?? "rpc error", error.data));
      } else {
        pending.resolve(message.result);
      }
    }
  }

  private async answerRequest(id: number | string, method: string, params: unknown): Promise<void> {
    try {
      if (!this.options.onRequest) {
        throw new JsonRpcError(-32601, `method not found: ${method}`);
      }
      const result = await this.options.onRequest(method, params, id);
      this.write({ id, result: result ?? null });
    } catch (error) {
      const code = error instanceof JsonRpcError ? error.code : -32603;
      const message = error instanceof Error ? error.message : String(error);
      this.write({ id, error: { code, message } });
    }
  }

  private teardown(reason: Error): void {
    if (this.closed) return;
    this.closed = true;
    for (const [id, pending] of this.pending) {
      this.pending.delete(id);
      pending.reject(reason);
    }
  }
}
