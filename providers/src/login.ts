import spawn from "cross-spawn";
import type { ChildProcess } from "node:child_process";
import type { ProviderId } from "../../protocol/src/index.js";
import type { ProviderLoginStatus } from "../../protocol/src/provider-login.js";
import { killProcessTree } from "./base.js";
import type { ProviderLoginAttempt } from "./types.js";

export class ProviderLoginError extends Error {}

/** How a CLI signs in: the command to run and how to read the page (and code) it prints. */
export interface CliLoginSpec {
  /** Name used in messages, for example "GitHub Copilot". */
  name: string;
  args: string[];
  /** Page and code once the CLI printed them; null while the output is still incomplete. */
  find(output: string): { uri: string; code: string | null } | null;
  /** The page hands back a code that must be written to the CLI's input. */
  pasteBack?: boolean;
  instructions: string;
}

export const LOGIN_TIMEOUT_MS = 10 * 60_000;
const STARTUP_TIMEOUT_MS = 30_000;
const ANSI = /\x1b\[[0-?]*[ -/]*[@-~]/g;

export function idleLoginStatus(provider: ProviderId, message: string): ProviderLoginStatus {
  return { state: "idle", provider, userCode: null, verificationUri: null, expiresAt: null, message };
}

/** Shared bookkeeping of one sign-in attempt: its status, its total time limit and the way it ends. */
export abstract class LoginAttempt implements ProviderLoginAttempt {
  protected current: ProviderLoginStatus;
  protected submitted = false;
  private readonly timers = new Set<NodeJS.Timeout>();

  constructor(protected readonly provider: ProviderId, message: string) {
    this.current = { state: "starting", provider, userCode: null, verificationUri: null, expiresAt: Date.now() + LOGIN_TIMEOUT_MS, message };
    this.after(LOGIN_TIMEOUT_MS, () => this.finish("expired", "The sign-in session timed out. Start again to continue."));
  }

  get active(): boolean { return this.current.state === "starting" || this.current.state === "waiting"; }
  status(): ProviderLoginStatus { return { ...this.current }; }
  abstract submitCode(code: string): void;
  cancel(): void { this.finish("cancelled", "Sign-in cancelled."); }

  protected after(ms: number, action: () => void): NodeJS.Timeout {
    const timer = setTimeout(action, ms);
    timer.unref();
    this.timers.add(timer);
    return timer;
  }

  protected clear(timer: NodeJS.Timeout): void { clearTimeout(timer); this.timers.delete(timer); }

  /** Shows the page (and code) to sign in with. */
  protected waiting(patch: Pick<ProviderLoginStatus, "userCode" | "verificationUri" | "message"> & { needsCode?: boolean }): void {
    this.current = { ...this.current, ...patch, state: "waiting" };
  }

  protected finish(state: ProviderLoginStatus["state"], message: string): void {
    if (!this.active) return;
    this.current = { state, provider: this.provider, userCode: null, verificationUri: null, expiresAt: null, message };
    for (const timer of this.timers) clearTimeout(timer);
    this.timers.clear();
    this.release();
  }

  /** Lets go of whatever the attempt holds (a process, a service). */
  protected abstract release(): void;
}

/** Signs in by running the provider's own CLI and reading the page and code it prints. */
export class CliLoginAttempt extends LoginAttempt {
  private child: ChildProcess | null = null;
  private output = "";

  constructor(provider: ProviderId, command: string, env: NodeJS.ProcessEnv, private readonly spec: CliLoginSpec) {
    super(provider, "Getting a sign-in link…");
    try {
      const child = spawn(command, spec.args, { windowsHide: true, stdio: [spec.pasteBack ? "pipe" : "ignore", "pipe", "pipe"], env: { ...env, NO_COLOR: "1", FORCE_COLOR: "0" } });
      this.child = child;
      child.stdin?.on("error", () => {});
      child.stdout?.on("data", chunk => this.receive(chunk));
      child.stderr?.on("data", chunk => this.receive(chunk));
      child.once("error", () => this.finish("error", `Could not run the ${spec.name} CLI. Install or update it, then try again.`));
      child.once("close", code => this.closed(code));
      this.startup = this.after(STARTUP_TIMEOUT_MS, () => this.finish("error", `The ${spec.name} CLI did not return a sign-in link. Update it and try again.`));
    } catch {
      this.finish("error", `Could not start the ${spec.name} CLI. Check the installation and try again.`);
    }
  }

  private startup: NodeJS.Timeout | null = null;

  private receive(chunk: Buffer): void {
    if (!this.active) return;
    this.output = (this.output + chunk.toString("utf8")).replace(ANSI, "").slice(-16_384);
    if (this.current.state !== "starting") return;
    const found = this.spec.find(this.output);
    if (!found) return;
    if (this.startup) this.clear(this.startup);
    this.waiting({ userCode: found.code, verificationUri: found.uri, needsCode: this.spec.pasteBack === true, message: this.spec.instructions });
  }

  private closed(code: number | null): void {
    if (!this.active) return;
    if (code === 0 && this.current.state === "waiting") return this.finish("success", `Signed in to ${this.spec.name}.`);
    const expired = /expired|timed?\s*out/i.test(this.output);
    this.finish(expired ? "expired" : "error", expired ? "The sign-in session expired. Start again to continue."
      : this.submitted ? "That code was not accepted. It may be incomplete or already used. Start again to get a new code."
      : `Sign-in did not finish. Update the ${this.spec.name} CLI and try again.`);
  }

  submitCode(code: string): void {
    if (!this.spec.pasteBack || this.current.state !== "waiting" || this.submitted) throw new ProviderLoginError("There is no sign-in waiting for a code.");
    const value = code.trim();
    if (!/^[\w#.~%+/=-]{4,512}$/.test(value)) throw new ProviderLoginError("That does not look like a sign-in code. Copy it again from the page.");
    this.submitted = true;
    this.current = { ...this.current, needsCode: false, message: "Checking the code…" };
    this.child?.stdin?.write(`${value}\n`);
  }

  protected release(): void {
    killProcessTree(this.child);
    this.child = null;
    this.output = "";
  }
}
