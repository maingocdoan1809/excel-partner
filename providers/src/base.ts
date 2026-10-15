import spawn from "cross-spawn";
import type {
  ProviderCapabilities,
  ProviderConnectionStatus,
  ProviderDetection,
  ProviderEvent,
  ProviderId,
} from "../../protocol/src/index.js";
import type { ProviderAdapter, ProviderRuntime, SendTurnInput, StartSessionContext } from "./types.js";

export const cliEnvironment = (): string => process.platform === "win32" ? "Windows" : "WSL/Linux";

/** Raised by adapters whose CLI handshake has not been verified yet. */
export class ProviderNotImplementedError extends Error {
  constructor(provider: ProviderId, operation: string) {
    super(
      `provider '${provider}' does not implement ${operation} yet; ` +
        `wire it against a real CLI install (see PLAN milestone 3)`,
    );
    this.name = "ProviderNotImplementedError";
  }
}

export interface CliProbe {
  command: string;
  versionArgs: string[];
  /** Command the user should run where the bridge is running to install the CLI. */
  installHint: string;
  /** Command the user should run where the bridge is running to log in. */
  loginHint: string;
}

export interface ProbeResult {
  installed: boolean;
  version: string | null;
  status: ProviderConnectionStatus;
  hint: string | null;
}

const DETECT_TTL_MS = 15_000;

export abstract class BaseCliAdapter implements ProviderAdapter {
  abstract readonly id: ProviderId;
  abstract readonly capabilities: ProviderCapabilities;

  private cached: { at: number; value: ProviderDetection } | null = null;

  constructor(protected readonly runtime: ProviderRuntime) {}

  protected abstract probe(): CliProbe;

  protected get env(): NodeJS.ProcessEnv {
    return process.env;
  }

  async detect(force = false): Promise<ProviderDetection> {
    if (!force && this.cached && Date.now() - this.cached.at < DETECT_TTL_MS) {
      return this.cached.value;
    }
    const { command, versionArgs, installHint, loginHint } = this.probe();
    const probed = await probeCli(command, versionArgs, this.env, installHint, loginHint);
    const minVersion = this.minVersion;
    if (probed.installed && probed.status === "ready" && minVersion && probed.version) {
      if (compareVersions(extractVersion(probed.version), minVersion) < 0) {
        probed.status = "error";
        probed.hint = `${this.id} CLI ${extractVersion(probed.version)} is older than the tested minimum ${minVersion}. Update it on ${cliEnvironment()}, then retry.`;
      }
    }
    if (probed.installed && probed.status === "ready") {
      const login = await this.checkLogin(command).catch((error) => ({
        status: "error" as ProviderConnectionStatus,
        hint: `Could not verify login: ${error instanceof Error ? error.message : String(error)}`,
      }));
      probed.status = login.status;
      probed.hint = login.hint;
    }
    const detection: ProviderDetection = {
      provider: this.id,
      installed: probed.installed,
      version: probed.version,
      status: probed.status,
      hint: probed.hint,
      capabilities: this.capabilities,
    };
    this.cached = { at: Date.now(), value: detection };
    this.runtime.logger.debug(`detect ${this.id}`, detection);
    return detection;
  }

  /** Oldest CLI version this adapter has been verified against, if enforced. */
  protected get minVersion(): string | null {
    return null;
  }

  /** Adapters override this to verify the CLI session is logged in. */
  protected async checkLogin(
    _command: string,
  ): Promise<{ status: ProviderConnectionStatus; hint: string | null }> {
    return { status: "ready", hint: null };
  }

  async connectionStatus(): Promise<ProviderConnectionStatus> {
    return (await this.detect()).status;
  }

  async startSession(_ctx: StartSessionContext): Promise<string> {
    throw new ProviderNotImplementedError(this.id, "startSession");
  }

  async *sendTurn(_input: SendTurnInput): AsyncIterable<ProviderEvent> {
    throw new ProviderNotImplementedError(this.id, "sendTurn");
  }

  async cancelTurn(_providerSessionId: string): Promise<void> {
    throw new ProviderNotImplementedError(this.id, "cancelTurn");
  }

  async closeSession(_providerSessionId: string): Promise<void> {
    /* nothing to release until a CLI process is spawned */
  }

  async dispose(): Promise<void> {
    /* no-op for the base adapter */
  }

  protected invalidateDetection(): void {
    this.cached = null;
  }
}

export interface ExecResult {
  code: number | null;
  stdout: string;
  stderr: string;
  errorCode?: string;
}

/** Run a short CLI command; never throws, a spawn failure yields code=null. */
export async function execCli(
  command: string,
  args: string[],
  env: NodeJS.ProcessEnv,
  timeoutMs = 10_000,
): Promise<ExecResult> {
  return new Promise((resolve) => {
    let child: ReturnType<typeof spawn>;
    try {
      // npm creates .cmd shims on Windows. cross-spawn resolves and quotes them
      // while preserving stdin/stdout for the ACP and app-server transports.
      child = spawn(command, args, { env, windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    } catch (error) {
      const err = error as NodeJS.ErrnoException;
      resolve({ code: null, stdout: "", stderr: err.message, errorCode: err.code });
      return;
    }
    let stdout = "";
    let stderr = "";
    let settled = false;
    let timedOut = false;
    const finish = (result: ExecResult) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(result);
    };
    const timer = setTimeout(() => { timedOut = true; child.kill(); }, timeoutMs);
    child.stdout?.on("data", (chunk: Buffer) => { stdout += chunk.toString(); });
    child.stderr?.on("data", (chunk: Buffer) => { stderr += chunk.toString(); });
    child.once("error", (error: NodeJS.ErrnoException) => finish({ code: null, stdout, stderr: stderr || error.message, errorCode: error.code }));
    child.once("close", (code) => finish({ code: timedOut ? null : code, stdout, stderr: timedOut ? "CLI timed out" : stderr, errorCode: timedOut ? "ETIMEDOUT" : undefined }));
  });
}

/**
 * Best-effort CLI probe. Never throws: a missing binary becomes
 * `not_installed` so the UI can show a hint instead of a hard failure.
 */
export async function probeCli(
  command: string,
  args: string[],
  env: NodeJS.ProcessEnv,
  installHint: string,
  loginHint: string,
): Promise<ProbeResult> {
  const result = await execCli(command, args, env, 8_000);
  if (result.code === 0) {
    const version = firstLine(result.stdout) || firstLine(result.stderr) || null;
    return { installed: true, version, status: "ready", hint: null };
  }
  const notFound = result.errorCode === "ENOENT" || /not found|no such file/i.test(result.stderr);
  return {
    installed: !notFound,
    version: null,
    status: notFound ? "not_installed" : "error",
    hint: notFound
      ? `Install it on ${cliEnvironment()}: ${installHint}`
      : `Could not run ${command} on ${cliEnvironment()}. ${result.stderr || `Exit code ${result.code}`}. Try \`${loginHint}\` there, then retry.`,
  };
}

function firstLine(value: string): string {
  return value.split(/\r?\n/, 1)[0]?.trim() ?? "";
}

/** Pull `1.2.3` out of strings like `codex-cli 0.162.1` or `2.1.294 (Claude Code)`. */
export function extractVersion(text: string): string {
  return /(\d+)\.(\d+)\.(\d+)/.exec(text)?.[0] ?? "0.0.0";
}

export function compareVersions(a: string, b: string): number {
  const pa = a.split(".").map(Number);
  const pb = b.split(".").map(Number);
  for (let i = 0; i < 3; i++) {
    const diff = (pa[i] ?? 0) - (pb[i] ?? 0);
    if (diff !== 0) return diff < 0 ? -1 : 1;
  }
  return 0;
}
