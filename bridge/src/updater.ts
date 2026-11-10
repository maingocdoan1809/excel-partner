import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

/** "0.1.9" < "0.1.19" < "0.2.0"; a release beats its own pre-release. Null when either side is not a version. */
export function compareVersions(a: string, b: string): number | null {
  const parse = (v: string) => {
    const m = /^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?$/.exec(v.trim());
    return m ? { n: [Number(m[1]), Number(m[2]), Number(m[3])], pre: m[4] ?? null } : null;
  };
  const x = parse(a);
  const y = parse(b);
  if (!x || !y) return null;
  for (let i = 0; i < 3; i++) if (x.n[i] !== y.n[i]) return x.n[i]! < y.n[i]! ? -1 : 1;
  if (x.pre === y.pre) return 0;
  if (x.pre === null) return 1;
  if (y.pre === null) return -1;
  return x.pre < y.pre ? -1 : 1;
}

export type UpdateState = "idle" | "running" | "done" | "failed";

export interface UpdateStatus {
  state: UpdateState;
  step?: string;
  message?: string;
  from?: string | null;
  to?: string | null;
  error?: string | null;
  startedAt?: string;
  finishedAt?: string | null;
}

export interface UpdateCheck {
  current: string;
  latest: string | null;
  updateAvailable: boolean;
  /** False when this installation cannot be replaced from the task pane (source checkout, npx, not Windows). */
  canUpdate: boolean;
  reason?: string;
  /** Why the latest version is unknown (offline, registry down). */
  checkError?: string;
  status: UpdateStatus;
  logFile: string;
}

export class UpdateError extends Error {
  constructor(readonly code: "busy" | "unsupported" | "running" | "failed", message: string) {
    super(message);
  }
}

export interface UpdateServiceOptions {
  /** The installed package folder (where package.json and bin/ live). */
  packageRoot: string;
  /** Excel Partner's data folder; holds update-status.json and update.log. */
  dataDir: string;
  /** Where the gateway keeps its token; present means it is set up and must be started again. */
  gatewayDir?: string;
  version?: string;
  platform?: NodeJS.Platform;
  /** What `npm install --global` is given; overridable for mirrors and tests. */
  spec?: string;
  fetchLatest?: () => Promise<string>;
  /** `npm root --global`; null when npm cannot be found. */
  globalRoot?: () => string | null;
  isPidAlive?: (pid: number) => boolean;
  /** Starts the detached runner; tests replace it. */
  launch?: (runnerPath: string, env: NodeJS.ProcessEnv, logFile: string) => number | undefined;
  now?: () => number;
  /** A "running" record older than this is treated as a crashed update. */
  staleAfterMs?: number;
  cacheMs?: number;
}

const REGISTRY_URL = "https://registry.npmjs.org/excel-partner/latest";

export async function fetchLatestFromRegistry(fetchImpl: typeof fetch = fetch, signal?: AbortSignal): Promise<string> {
  const response = await fetchImpl(REGISTRY_URL, { headers: { accept: "application/json" }, signal: signal ?? AbortSignal.timeout(8000) });
  if (!response.ok) throw new Error(`the npm registry answered HTTP ${response.status}`);
  const body = (await response.json()) as { version?: unknown };
  if (typeof body.version !== "string" || compareVersions(body.version, "0.0.0") === null) throw new Error("the npm registry sent no version");
  return body.version;
}

export class UpdateService {
  private readonly opts: Required<Pick<UpdateServiceOptions, "platform" | "spec" | "now" | "staleAfterMs" | "cacheMs">> & UpdateServiceOptions;
  private cached: { at: number; latest: string | null; error?: string } | null = null;
  private globalRootCache: string | null | undefined;

  constructor(options: UpdateServiceOptions) {
    this.opts = {
      platform: process.platform,
      spec: process.env.EXCEL_PARTNER_UPDATE_SPEC || "excel-partner@latest",
      now: Date.now,
      staleAfterMs: 15 * 60_000,
      cacheMs: 10 * 60_000,
      ...options,
    };
  }

  get version(): string {
    if (this.opts.version) return this.opts.version;
    try {
      return JSON.parse(fs.readFileSync(path.join(this.opts.packageRoot, "package.json"), "utf8")).version as string;
    } catch {
      return "0.0.0";
    }
  }

  get logFile(): string {
    return path.join(this.opts.dataDir, "update.log");
  }

  private statusFile(): string {
    return path.join(this.opts.dataDir, "update-status.json");
  }

  /** The record of the last or running update. A running record whose runner is gone is reported as failed. */
  status(): UpdateStatus {
    let raw: (UpdateStatus & { pid?: number }) | null = null;
    try {
      raw = JSON.parse(fs.readFileSync(this.statusFile(), "utf8"));
    } catch {
      return { state: "idle" };
    }
    if (!raw || typeof raw.state !== "string") return { state: "idle" };
    const { pid, ...status } = raw;
    if (status.state === "running") {
      const started = status.startedAt ? Date.parse(status.startedAt) : NaN;
      // The runner records its own pid a moment after it starts; until then trust the start time.
      const alive = typeof pid === "number" ? (this.opts.isPidAlive ?? isPidAlive)(pid) : Number.isFinite(started) && this.opts.now() - started < 20_000;
      const stale = Number.isFinite(started) && this.opts.now() - started > this.opts.staleAfterMs;
      if (!alive || stale) return { ...status, state: "failed", error: "The update stopped before it finished.", message: "The update did not finish" };
    }
    return status as UpdateStatus;
  }

  private globalRoot(): string | null {
    if (this.globalRootCache !== undefined) return this.globalRootCache;
    let found: string | null = null;
    try {
      found = this.opts.globalRoot ? this.opts.globalRoot() : npmGlobalRoot();
    } catch {
      found = null;
    }
    this.globalRootCache = found;
    return found;
  }

  /** Why this installation cannot update itself, or null when it can. */
  unsupportedReason(): string | null {
    if (this.opts.platform !== "win32") return "Updating from the task pane works on Windows with Excel desktop.";
    const root = this.globalRoot();
    if (!root) return "npm was not found. Install Node.js, or run excel-partner update from a terminal.";
    if (path.resolve(root, "excel-partner").toLowerCase() !== path.resolve(this.opts.packageRoot).toLowerCase()) {
      return "This copy was not installed with npm install -g, so it cannot replace itself. From a source checkout run scripts\\update.ps1.";
    }
    return null;
  }

  async check(refresh = false): Promise<UpdateCheck> {
    const now = this.opts.now();
    if (refresh || !this.cached || now - this.cached.at > this.opts.cacheMs) {
      try {
        this.cached = { at: now, latest: await (this.opts.fetchLatest ?? (() => fetchLatestFromRegistry()))() };
      } catch (error) {
        this.cached = { at: now, latest: null, error: error instanceof Error ? error.message : String(error) };
      }
    }
    const current = this.version;
    const latest = this.cached.latest;
    const newer = latest !== null && compareVersions(latest, current) === 1;
    const reason = this.unsupportedReason();
    return {
      current,
      latest,
      updateAvailable: newer,
      canUpdate: reason === null,
      ...(reason ? { reason } : {}),
      ...(this.cached.error ? { checkError: this.cached.error } : {}),
      status: this.status(),
      logFile: this.logFile,
    };
  }

  /** Start the update in a separate process; it outlives this bridge, which it stops. */
  start(options: { busy: boolean }): UpdateStatus {
    const reason = this.unsupportedReason();
    if (reason) throw new UpdateError("unsupported", reason);
    if (this.status().state === "running") throw new UpdateError("running", "An update is already running.");
    if (options.busy) throw new UpdateError("busy", "The assistant is still working. Wait for it to finish, or stop it, then update.");

    const runnerDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-partner-update-"));
    for (const file of ["update-runner.mjs", "update.mjs"]) fs.copyFileSync(path.join(this.opts.packageRoot, "bin", file), path.join(runnerDir, file));
    // Lets another updater identify this runner's installation without guessing from its temporary folder name.
    fs.writeFileSync(path.join(runnerDir, "target.json"), JSON.stringify({ root: this.opts.packageRoot }));
    fs.mkdirSync(this.opts.dataDir, { recursive: true });
    const gatewaySetUp = !!this.opts.gatewayDir && fs.existsSync(path.join(this.opts.gatewayDir, "token.txt"));
    const env: NodeJS.ProcessEnv = {
      ...process.env,
      EP_ROOT: this.opts.packageRoot,
      EP_DATA: this.opts.dataDir,
      EP_SPEC: this.opts.spec,
      EP_FROM: this.version,
      EP_GATEWAY: gatewaySetUp ? "1" : "0",
    };
    const started: UpdateStatus = {
      state: "running", step: "preparing", message: "Getting ready", from: this.version, to: null, error: null,
      startedAt: new Date(this.opts.now()).toISOString(), finishedAt: null,
    };
    fs.writeFileSync(this.statusFile(), JSON.stringify(started, null, 2));
    try {
      (this.opts.launch ?? launchDetached)(path.join(runnerDir, "update-runner.mjs"), env, this.logFile);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      fs.writeFileSync(this.statusFile(), JSON.stringify({ ...started, state: "failed", error: message, message: "The update did not start", finishedAt: new Date(this.opts.now()).toISOString() }, null, 2));
      throw new UpdateError("failed", `The update could not start: ${message}`);
    }
    return started;
  }
}

function isPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function npmGlobalRoot(): string | null {
  // npm.cmd needs the Windows shell; the arguments are fixed.
  const result = spawnSync("npm.cmd", ["root", "--global"], { shell: true, encoding: "utf8", windowsHide: true, timeout: 15_000 });
  const out = result.status === 0 ? String(result.stdout).trim() : "";
  return out || null;
}

function launchDetached(runnerPath: string, env: NodeJS.ProcessEnv, logFile: string): number | undefined {
  const log = fs.openSync(logFile, "a");
  fs.writeSync(log, `\n--- update started ${new Date().toISOString()} ---\n`);
  const child = spawn(process.execPath, [runnerPath], { cwd: env.EP_DATA, env, detached: true, windowsHide: true, stdio: ["ignore", log, log] });
  child.unref();
  fs.closeSync(log);
  return child.pid;
}
