import type { ProviderId } from "./providers.js";

/** A question a provider asks before it can start a sign-in (for example the GitHub deployment type). */
export interface LoginPrompt {
  key: string;
  type: "text" | "select";
  message: string;
  placeholder?: string;
  options?: { label: string; value: string; hint?: string }[];
  /** Ask only while another answer matches. */
  when?: { key: string; op: "eq" | "neq"; value: string };
}

export interface LoginMethod {
  /** Position in the provider's list for the account; the sign-in refers to it. */
  index: number;
  /** `oauth` opens a page (and may show a code); `api` takes a key. */
  type: "oauth" | "api";
  label: string;
  prompts: LoginPrompt[];
}

/** One account a provider can use (login kind `account`), with the ways it can be signed in. */
export interface LoginAccount {
  id: string;
  name: string;
  /** Has credentials already. */
  connected: boolean;
  methods: LoginMethod[];
}

/** What the person picked in the pane. `key` is passed straight to the provider and kept nowhere else. */
export interface LoginChoice {
  accountId: string;
  method: number;
  inputs?: Record<string, string>;
  key?: string;
}

export interface ProviderLoginStatus {
  state: "idle" | "starting" | "waiting" | "success" | "cancelled" | "expired" | "error";
  /** Provider this attempt belongs to. */
  provider?: ProviderId;
  userCode: string | null;
  /** HTTPS page to open to finish signing in. */
  verificationUri: string | null;
  /** True while the page returns a code that the user must paste back. */
  needsCode?: boolean;
  expiresAt: number | null;
  message: string;
}
