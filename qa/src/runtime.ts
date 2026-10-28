import fs from "node:fs";
import path from "node:path";
import spawn from "cross-spawn";
import { AUTOMATION_DIR, allowedOriginsOf, requireEnvironment, requireProfile, type QaConfig } from "./config.js";
import { authFile } from "./session.js";
import { QaError, redact, scrubbedEnv } from "./util.js";
import type { Credentials } from "./vault.js";

export const nodeBin = (): string => process.env.QA_NODE_BIN || process.execPath;

/** The only variables a test or login process receives besides a minimal OS set: they describe where, never who. */
export function qaEnv(projectDir: string, config: QaConfig, environmentName: string | undefined, extra: Record<string, string | undefined> = {}): NodeJS.ProcessEnv {
  const environment = requireEnvironment(config, environmentName);
  return scrubbedEnv({
    QA_PROJECT_DIR: projectDir,
    QA_ENVIRONMENT: environment.name,
    QA_BASE_URL: environment.baseUrl,
    QA_ALLOWED_ORIGINS: allowedOriginsOf(config).join(","),
    QA_AUTH_DIR: path.join(projectDir, "auth", environment.name),
    QA_TIMEOUT_MS: String(config.run.timeoutMs),
    QA_WORKERS: String(config.run.workers),
    QA_RETRIES: String(config.run.retries),
    QA_IGNORE_HTTPS_ERRORS: config.environments[environment.name]!.ignoreHttpsErrors ? "1" : undefined,
    ...extra,
  });
}

export function playwrightCli(projectDir: string): string {
  const cli = path.join(projectDir, "node_modules", "@playwright", "test", "cli.js");
  if (!fs.existsSync(cli)) throw new QaError(`Playwright is not installed in this project. Run: cd "${projectDir}" && npm install && npx playwright install chromium (or: excel-partner qa doctor --fix)`, "no_playwright");
  return cli;
}

export interface ProcessResult { code: number | null; signal: string | null; stdout: string; stderr: string; timedOut: boolean; cancelled: boolean }

export interface SpawnOptions { cwd: string; env: NodeJS.ProcessEnv; args: string[]; stdin?: string; timeoutMs: number; signal?: AbortSignal; inheritStderr?: boolean; inheritStdout?: boolean; onLine?: (line: string) => void }

/** Run Node with a scrubbed environment, a time limit and tree-kill on cancel. Output is capped and redacted. */
export function runNode(options: SpawnOptions): Promise<ProcessResult> {
  return new Promise((resolve, reject) => {
    if (options.signal?.aborted) { resolve({ code: null, signal: null, stdout: "", stderr: "", timedOut: false, cancelled: true }); return; }
    const child = spawn(nodeBin(), options.args, {
      cwd: options.cwd, env: options.env, windowsHide: !options.inheritStderr, shell: false,
      stdio: [options.stdin === undefined ? "ignore" : "pipe", options.inheritStdout ? "inherit" : "pipe", options.inheritStderr ? "inherit" : "pipe"],
    });
    let stdout = "", stderr = "", timedOut = false, cancelled = false, stopped = false;
    const cap = (text: string, add: string) => (text + add).slice(-64_000);
    const stop = () => {
      if (stopped || !child.pid) return;
      stopped = true;
      if (process.platform === "win32") {
        const killer = spawn("taskkill", ["/pid", String(child.pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
        killer.once("error", () => child.kill("SIGKILL"));
      } else child.kill("SIGKILL");
    };
    let pending = "";
    child.stdout?.on("data", (chunk: Buffer) => {
      const text = chunk.toString("utf8");
      stdout = cap(stdout, text);
      if (options.onLine) {
        pending += text;
        const lines = pending.split(/\r?\n/);
        pending = lines.pop() ?? "";
        for (const line of lines) options.onLine(redact(line));
      }
    });
    child.stderr?.on("data", (chunk: Buffer) => { stderr = cap(stderr, chunk.toString("utf8")); });
    const timer = setTimeout(() => { timedOut = true; stop(); }, options.timeoutMs);
    const abort = () => { cancelled = true; stop(); };
    options.signal?.addEventListener("abort", abort, { once: true });
    const cleanup = () => { clearTimeout(timer); options.signal?.removeEventListener("abort", abort); };
    child.once("error", error => { cleanup(); reject(new QaError(`Could not start Node: ${error.message}`, "spawn_error")); });
    child.once("close", (code, signal) => {
      cleanup();
      if (options.onLine && pending) options.onLine(redact(pending));
      resolve({ code, signal, stdout: redact(stdout), stderr: redact(stderr), timedOut, cancelled });
    });
    if (options.stdin !== undefined) { child.stdin?.on("error", () => undefined); child.stdin?.end(options.stdin); }
  });
}

export interface LoginOptions {
  projectDir: string; config: QaConfig; environment?: string; profile: string;
  mode: "interactive" | "headless"; credentials?: Credentials; signal?: AbortSignal; timeoutMs?: number;
}

/** Sign in through the harness script and save the browser session to auth/<env>/<profile>.json. */
export async function runLogin(options: LoginOptions): Promise<void> {
  const profile = requireProfile(options.config, options.profile);
  const environment = requireEnvironment(options.config, options.environment);
  playwrightCli(options.projectDir);
  const script = path.join(options.projectDir, AUTOMATION_DIR, "_harness", "login.js");
  if (!fs.existsSync(script)) throw new QaError("The harness is not installed. Run: excel-partner qa upgrade", "no_harness");
  const timeoutMs = options.timeoutMs ?? (options.mode === "interactive" ? 330_000 : 90_000);
  const result = await runNode({
    cwd: options.projectDir, args: [script], timeoutMs, signal: options.signal,
    stdin: options.mode === "headless" ? JSON.stringify(options.credentials) : undefined,
    inheritStderr: options.mode === "interactive", inheritStdout: false,
    env: qaEnv(options.projectDir, options.config, environment.name, {
      QA_LOGIN_MODE: options.mode, QA_LOGIN_PATH: profile.login.path, QA_USER_SELECTOR: profile.login.userSelector,
      QA_PASSWORD_SELECTOR: profile.login.passwordSelector, QA_SUBMIT_SELECTOR: profile.login.submitSelector,
      QA_LOGIN_SUCCESS_URL: profile.login.successUrl, QA_CHECK_TYPE: profile.check.type, QA_CHECK_PATH: profile.check.path, QA_LOGIN_TIMEOUT_MS: String(timeoutMs - 20_000),
      QA_STORAGE_OUT: authFile(options.projectDir, environment.name, options.profile),
    }),
  });
  if (result.cancelled) throw new QaError("Sign-in was cancelled.", "cancelled");
  if (result.timedOut) throw new QaError("Sign-in timed out.", "timeout");
  if (result.code !== 0) throw new QaError((result.stderr.trim().split("\n").pop() || `Sign-in failed (exit ${result.code}).`), "login_failed");
}
