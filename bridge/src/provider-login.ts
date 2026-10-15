import spawn from "cross-spawn";
import type { ChildProcess } from "node:child_process";
import type { ProviderLoginStatus } from "../../protocol/src/provider-login.js";

interface LoginAttempt {
  status: ProviderLoginStatus;
  child: ChildProcess | null;
  timer: NodeJS.Timeout | null;
  startupTimer: NodeJS.Timeout | null;
  output: string;
}

export class ProviderLoginError extends Error {}

export class ProviderLoginManager {
  private readonly attempts = new Map<string, LoginAttempt>();

  constructor(private readonly command = process.env.COPILOT_BIN ?? "copilot") {}

  status(sessionId: string): ProviderLoginStatus {
    return { ...(this.attempts.get(sessionId)?.status ?? {
      state: "idle", userCode: null, verificationUri: null, expiresAt: null,
      message: "Connect your GitHub Copilot account to get started.",
    }) };
  }

  start(sessionId: string): ProviderLoginStatus {
    for (const [owner, attempt] of this.attempts) {
      if (attempt.status.state !== "starting" && attempt.status.state !== "waiting") continue;
      if (owner === sessionId) return this.status(sessionId);
      throw new ProviderLoginError("Another Excel window is signing in. Finish or cancel that session first.");
    }
    const attempt: LoginAttempt = {
      status: { state: "starting", userCode: null, verificationUri: null, expiresAt: Date.now() + 10 * 60_000, message: "Getting a device code…" },
      child: null, timer: null, startupTimer: null, output: "",
    };
    this.attempts.set(sessionId, attempt);
    try {
      const child = spawn(this.command, ["login", "--device-code"], {
        windowsHide: true, stdio: ["ignore", "pipe", "pipe"],
        env: { ...process.env, NO_COLOR: "1", FORCE_COLOR: "0" },
      });
      attempt.child = child;
      const receive = (chunk: Buffer) => {
        if (!this.isActive(attempt)) return;
        attempt.output = (attempt.output + chunk.toString("utf8")).replace(/\x1b\[[0-?]*[ -/]*[@-~]/g, "").slice(-16_384);
        const code = /(?:enter(?:\s+(?:the|one-time))?\s+code|(?:one-time|device|user)\s+code|code)\s*(?:is\s*)?[:=]?\s*([A-Z0-9]{4}-[A-Z0-9]{4})\b/i.exec(attempt.output);
        if (code && /https:\/\/github\.com\/login\/device\b/.test(attempt.output)) {
          if (attempt.startupTimer) clearTimeout(attempt.startupTimer);
          attempt.startupTimer = null;
          attempt.status.state = "waiting";
          attempt.status.userCode = code[1]!.toUpperCase();
          attempt.status.verificationUri = "https://github.com/login/device";
          attempt.status.message = "Enter the code on GitHub. Excel Partner will update when you finish signing in.";
        }
      };
      child.stdout?.on("data", receive);
      child.stderr?.on("data", receive);
      child.once("error", () => this.finish(attempt, "error", "Could not run Copilot CLI. Install or update @github/copilot on the bridge machine, then try again."));
      child.once("close", (code) => {
        if (!this.isActive(attempt)) return;
        if (code === 0 && attempt.status.userCode) {
          this.finish(attempt, "success", "Signed in to GitHub Copilot.");
        } else {
          const expired = /expired|timed?\s*out/i.test(attempt.output);
          this.finish(attempt, expired ? "expired" : "error", expired
            ? "The sign-in session expired. Get a new code to continue."
            : "Sign-in did not finish. Update Copilot CLI and try again, or run copilot login --device-code in a terminal.");
        }
      });
      attempt.timer = setTimeout(() => this.finish(attempt, "expired", "The sign-in session timed out. Get a new code to continue."), 10 * 60_000);
      attempt.timer.unref();
      attempt.startupTimer = setTimeout(() => this.finish(attempt, "error", "Copilot CLI did not return a device code. Update the CLI or run copilot login --device-code in a terminal, then try again."), 30_000);
      attempt.startupTimer.unref();
    } catch {
      this.finish(attempt, "error", "Could not start Copilot CLI. Check the installation and try again.");
    }
    return this.status(sessionId);
  }

  cancel(sessionId: string): ProviderLoginStatus {
    const attempt = this.attempts.get(sessionId);
    if (attempt && this.isActive(attempt)) this.finish(attempt, "cancelled", "Sign-in cancelled.");
    return this.status(sessionId);
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
    attempt.status = { state, userCode: null, verificationUri: null, expiresAt: null, message };
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
