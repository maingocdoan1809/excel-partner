import spawn from "cross-spawn";
import type { ChildProcess } from "node:child_process";
import type { ProviderId } from "../../protocol/src/providers.js";
import { APP_LOGIN_PROVIDERS, type ProviderLoginStatus } from "../../protocol/src/provider-login.js";
import { activateProviderAfterLogin, providerExecution } from "../../providers/src/profiled.js";

interface LoginSpec {
  name: string;
  args: string[];
  /** Command a user can run by hand when the CLI does not behave. */
  manual: string;
  /** Page and code once the CLI printed them; null while the output is still incomplete. */
  find(output: string): { uri: string; code: string | null } | null;
  /** The page hands back a code that must be written to the CLI's input. */
  pasteBack?: boolean;
  instructions: string;
}

const LOGIN: Partial<Record<ProviderId, LoginSpec>> = {
  copilot: {
    name: "GitHub Copilot", args: ["login", "--device-code"], manual: "copilot login --device-code",
    find(output) {
      const code = /(?:enter(?:\s+(?:the|one-time))?\s+code|(?:one-time|device|user)\s+code|code)\s*(?:is\s*)?[:=]?\s*([A-Z0-9]{4}-[A-Z0-9]{4})\b/i.exec(output);
      return code && /https:\/\/github\.com\/login\/device\b/.test(output) ? { uri: "https://github.com/login/device", code: code[1]!.toUpperCase() } : null;
    },
    instructions: "Enter the code on GitHub. Office Partner will update when you finish signing in.",
  },
  codex: {
    name: "Codex", args: ["login", "--device-auth"], manual: "codex login --device-auth",
    find(output) {
      const code = /one-time code[^\n]*\n\s*([A-Z0-9]{4}-[A-Z0-9]{4,6})\b/i.exec(output);
      return code && /https:\/\/auth\.openai\.com\/codex\/device\b/.test(output) ? { uri: "https://auth.openai.com/codex/device", code: code[1]!.toUpperCase() } : null;
    },
    instructions: "Sign in on the OpenAI page and enter the code. Office Partner will update when you finish.",
  },
  claude: {
    name: "Claude", args: ["auth", "login"], manual: "claude auth login", pasteBack: true,
    find(output) {
      const link = /(https:\/\/claude\.(?:com|ai)\/\S+)\s/.exec(output);
      return link ? { uri: link[1]!, code: null } : null;
    },
    instructions: "Sign in on the Claude page, then paste the code it shows below.",
  },
};

interface LoginAttempt {
  provider: ProviderId;
  spec: LoginSpec;
  status: ProviderLoginStatus;
  child: ChildProcess | null;
  timer: NodeJS.Timeout | null;
  startupTimer: NodeJS.Timeout | null;
  output: string;
  submitted: boolean;
}

export class ProviderLoginError extends Error {}

export class ProviderLoginManager {
  private readonly attempts = new Map<string, LoginAttempt>();

  /** `command` replaces the provider CLI (tests). */
  constructor(private readonly command?: string) {}

  get isBusy() { return [...this.attempts.values()].some(attempt => this.isActive(attempt)); }

  status(sessionId: string, provider: ProviderId = "copilot"): ProviderLoginStatus {
    const attempt = this.attempts.get(sessionId);
    if (attempt?.provider === provider) return { ...attempt.status };
    return {
      state: "idle", provider, userCode: null, verificationUri: null, expiresAt: null,
      message: LOGIN[provider] ? `Connect your ${LOGIN[provider]!.name} account to get started.` : "Sign in with the command shown below.",
    };
  }

  start(sessionId: string, provider: ProviderId = "copilot"): ProviderLoginStatus {
    const spec = LOGIN[provider];
    if (!spec || !APP_LOGIN_PROVIDERS.includes(provider)) throw new ProviderLoginError("This provider signs in from a terminal.");
    for (const [owner, attempt] of this.attempts) {
      if (!this.isActive(attempt)) continue;
      if (owner !== sessionId) throw new ProviderLoginError("Another Excel window is signing in. Finish or cancel that session first.");
      if (attempt.provider === provider) return this.status(sessionId, provider);
      this.finish(attempt, "cancelled", "Sign-in cancelled.");
    }
    const attempt: LoginAttempt = {
      provider, spec, child: null, timer: null, startupTimer: null, output: "", submitted: false,
      status: { state: "starting", provider, userCode: null, verificationUri: null, expiresAt: Date.now() + 10 * 60_000, message: "Getting a sign-in link…" },
    };
    this.attempts.set(sessionId, attempt);
    try {
      const execution = this.command ? { command: this.command, env: process.env } : providerExecution(provider, process.env, true);
      const child = spawn(execution.command, spec.args, {
        windowsHide: true, stdio: [spec.pasteBack ? "pipe" : "ignore", "pipe", "pipe"],
        env: { ...execution.env, NO_COLOR: "1", FORCE_COLOR: "0" },
      });
      attempt.child = child;
      child.stdin?.on("error", () => {});
      const receive = (chunk: Buffer) => {
        if (!this.isActive(attempt)) return;
        attempt.output = (attempt.output + chunk.toString("utf8")).replace(/\x1b\[[0-?]*[ -/]*[@-~]/g, "").slice(-16_384);
        if (attempt.status.state !== "starting") return;
        const found = spec.find(attempt.output);
        if (!found) return;
        if (attempt.startupTimer) clearTimeout(attempt.startupTimer);
        attempt.startupTimer = null;
        attempt.status = { ...attempt.status, state: "waiting", userCode: found.code, verificationUri: found.uri, needsCode: spec.pasteBack === true, message: spec.instructions };
      };
      child.stdout?.on("data", receive);
      child.stderr?.on("data", receive);
      child.once("error", () => this.finish(attempt, "error", `Could not run the ${spec.name} CLI. Install or update it on the bridge machine, then try again.`));
      child.once("close", (code) => {
        if (!this.isActive(attempt)) return;
        if (code === 0 && attempt.status.state === "waiting") {
          if (this.command) this.finish(attempt, "success", `Signed in to ${spec.name}.`);
          else {
            attempt.status = { ...attempt.status, needsCode: false, message: "Verifying the signed-in provider runtime…" };
            void activateProviderAfterLogin(provider, process.env, () => this.isActive(attempt)).then(() => this.finish(attempt, "success", `Signed in to ${spec.name}.`),
              () => this.finish(attempt, "error", "Sign-in finished, but runtime verification failed. Run office-partner doctor and retry provider installation."));
          }
        } else {
          const expired = /expired|timed?\s*out/i.test(attempt.output);
          this.finish(attempt, expired ? "expired" : "error", expired
            ? "The sign-in session expired. Start again to continue."
            : attempt.submitted ? "That code was not accepted. It may be incomplete or already used. Start again to get a new code."
            : `Sign-in did not finish. Update the ${spec.name} CLI and try again, or run ${spec.manual} in a terminal.`);
        }
      });
      attempt.timer = setTimeout(() => this.finish(attempt, "expired", "The sign-in session timed out. Start again to continue."), 10 * 60_000);
      attempt.timer.unref();
      attempt.startupTimer = setTimeout(() => this.finish(attempt, "error", `The ${spec.name} CLI did not return a sign-in link. Update the CLI or run ${spec.manual} in a terminal, then try again.`), 30_000);
      attempt.startupTimer.unref();
    } catch {
      this.finish(attempt, "error", `Could not start the ${spec.name} CLI. Check the installation and try again.`);
    }
    return this.status(sessionId, provider);
  }

  /** Writes the code shown by the sign-in page to the CLI; only for providers that hand one back. */
  submitCode(sessionId: string, code: string): ProviderLoginStatus {
    const attempt = this.attempts.get(sessionId);
    if (!attempt || !this.isActive(attempt) || !attempt.spec.pasteBack || attempt.status.state !== "waiting" || attempt.submitted)
      throw new ProviderLoginError("There is no sign-in waiting for a code.");
    const value = code.trim();
    if (!/^[\w#.~%+/=-]{4,512}$/.test(value)) throw new ProviderLoginError("That does not look like a sign-in code. Copy it again from the page.");
    attempt.submitted = true;
    attempt.status = { ...attempt.status, needsCode: false, message: "Checking the code…" };
    attempt.child?.stdin?.write(`${value}\n`);
    return this.status(sessionId, attempt.provider);
  }

  cancel(sessionId: string): ProviderLoginStatus {
    const attempt = this.attempts.get(sessionId);
    if (attempt && this.isActive(attempt)) this.finish(attempt, "cancelled", "Sign-in cancelled.");
    return this.status(sessionId, attempt?.provider);
  }

  forget(sessionId: string): void {
    this.cancel(sessionId);
    this.attempts.delete(sessionId);
  }

  dispose(): void {
    for (const sessionId of this.attempts.keys()) this.forget(sessionId);
  }

  private isActive(attempt: LoginAttempt): boolean {
    return attempt.status.state === "starting" || attempt.status.state === "waiting";
  }

  private finish(attempt: LoginAttempt, state: ProviderLoginStatus["state"], message: string): void {
    if (!this.isActive(attempt)) return;
    attempt.status = { state, provider: attempt.provider, userCode: null, verificationUri: null, expiresAt: null, message };
    attempt.output = "";
    if (attempt.timer) clearTimeout(attempt.timer);
    if (attempt.startupTimer) clearTimeout(attempt.startupTimer);
    attempt.timer = null;
    attempt.startupTimer = null;
    const child = attempt.child;
    if (child?.exitCode === null && child.signalCode === null) {
      if (process.platform === "win32" && child.pid) {
        const killer = spawn("taskkill", ["/pid", String(child.pid), "/t", "/f"], { windowsHide: true, stdio: "ignore" });
        killer.on("error", () => child.kill());
      } else child.kill();
    }
    attempt.child = null;
  }
}
