import type { ProviderId } from "./providers.js";

/** Providers whose CLI can sign in from the pane: it prints a link (and a code) that the pane shows. */
export const APP_LOGIN_PROVIDERS: readonly ProviderId[] = ["copilot", "codex", "claude"];

export interface ProviderLoginStatus {
  state: "idle" | "starting" | "waiting" | "success" | "cancelled" | "expired" | "error";
  /** Provider this attempt belongs to. */
  provider?: ProviderId;
  userCode: string | null;
  /** HTTPS page to open to finish signing in. */
  verificationUri: string | null;
  /** True while the page returns a code that the user must paste back (Claude). */
  needsCode?: boolean;
  expiresAt: number | null;
  message: string;
}
