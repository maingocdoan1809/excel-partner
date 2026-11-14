import spawn from "cross-spawn";
import { randomBytes } from "node:crypto";
import type { ChildProcess } from "node:child_process";
import os from "node:os";
import { setTimeout as delay } from "node:timers/promises";
import type { LoginAccount, LoginChoice, LoginMethod, LoginPrompt } from "../../../protocol/src/provider-login.js";
import { killProcessTree } from "../base.js";
import { LoginAttempt, ProviderLoginError } from "../login.js";
import { reserveOpencodePort } from "./questions.js";

export interface OpenCodeAuthorization {
  url: string;
  /** `auto`: OpenCode waits for the page to finish. `code`: the person pastes back a code. */
  method: "auto" | "code";
  instructions: string;
}

/** The part of OpenCode's sign-in the pane needs. A person never types a command: the pane drives this. */
export interface OpenCodeAuthApi {
  readonly closed: boolean;
  catalog(): Promise<LoginAccount[]>;
  authorize(accountId: string, method: number, inputs: Record<string, string>): Promise<OpenCodeAuthorization>;
  /** Resolves once OpenCode has saved the account; `code` is the pasted code for a `code` method. */
  complete(accountId: string, method: number, code?: string): Promise<void>;
  setKey(accountId: string, key: string, metadata?: Record<string, string>): Promise<void>;
  close(): void;
}

export class OpenCodeAuthError extends ProviderLoginError {}

/** The service is shut down after this long without use. */
const IDLE_MS = 10 * 60_000;
/** The code OpenCode prints for a device sign-in, for example `Enter code: ABCD-EFGH`. */
const DEVICE_CODE = /code[:\s]+([A-Z0-9]{4,}(?:-[A-Z0-9]{3,})+)\b/i;
const SAFE_ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,79}$/;
/** Listed first, in this order, because most people come for one of them. */
const POPULAR = ["openai", "github-copilot", "anthropic", "google", "openrouter", "opencode"];

interface Options {
  spawn?: typeof spawn;
  fetch?: typeof fetch;
  port?: () => Promise<number>;
  startupMs?: number;
  idleMs?: number;
}

const str = (value: unknown, max = 400): string | undefined => typeof value === "string" && value.length <= max ? value : undefined;

function prompts(value: unknown): LoginPrompt[] {
  if (!Array.isArray(value)) return [];
  const result: LoginPrompt[] = [];
  for (const item of value as Record<string, unknown>[]) {
    const key = str(item?.key, 80), message = str(item?.message);
    if (!key || !message || (item.type !== "text" && item.type !== "select")) continue;
    const prompt: LoginPrompt = { key, type: item.type, message };
    const placeholder = str(item.placeholder);
    if (placeholder) prompt.placeholder = placeholder;
    if (Array.isArray(item.options)) {
      prompt.options = (item.options as Record<string, unknown>[]).flatMap(option => {
        const label = str(option?.label), optionValue = str(option?.value);
        return label && optionValue !== undefined ? [{ label, value: optionValue, ...(str(option.hint) ? { hint: str(option.hint)! } : {}) }] : [];
      });
    }
    const when = item.when as Record<string, unknown> | undefined;
    if (when && str(when.key) && (when.op === "eq" || when.op === "neq") && str(when.value) !== undefined) {
      prompt.when = { key: str(when.key)!, op: when.op, value: str(when.value)! };
    }
    result.push(prompt);
  }
  return result;
}

/**
 * Talks to a short-lived `opencode serve` bound to 127.0.0.1 with a random password and port.
 * It is OpenCode's own sign-in API: credentials are written by OpenCode into its own store.
 */
export class OpenCodeAuthServer implements OpenCodeAuthApi {
  private starting: Promise<void> | null = null;
  private child: ChildProcess | null = null;
  private base = "";
  private authorization = "";
  private idle: NodeJS.Timeout | null = null;
  private readonly abort = new AbortController();
  private isClosed = false;

  constructor(private readonly execution: { command: string; env: NodeJS.ProcessEnv }, private readonly options: Options = {}) {}

  get closed(): boolean { return this.isClosed; }

  private async request<T>(method: "GET" | "POST" | "PUT", path: string, body?: unknown): Promise<T> {
    const response = await (this.options.fetch ?? fetch)(`${this.base}${path}`, {
      method, signal: this.abort.signal,
      headers: { authorization: this.authorization, ...(body === undefined ? {} : { "content-type": "application/json" }) },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    });
    const text = await response.text();
    let data: unknown;
    try { data = text ? JSON.parse(text) : null; } catch { data = null; }
    if (!response.ok) {
      const detail = (data as { data?: { message?: unknown } } | null)?.data?.message;
      throw new OpenCodeAuthError(typeof detail === "string" && detail.length < 300 ? detail : `OpenCode refused the request (${response.status}).`);
    }
    return data as T;
  }

  /** Starts the private server on first use and keeps it for as long as it is being used. */
  private async start(): Promise<void> {
    if (this.isClosed) throw new OpenCodeAuthError("Sign-in was cancelled.");
    if (this.idle) clearTimeout(this.idle);
    this.idle = setTimeout(() => this.close(), this.options.idleMs ?? IDLE_MS);
    this.idle.unref();
    this.starting ??= this.launch();
    this.starting.catch(() => this.close());
    return this.starting;
  }

  private async launch(): Promise<void> {
    const port = await (this.options.port ?? reserveOpencodePort)();
    const password = randomBytes(18).toString("hex");
    this.base = `http://127.0.0.1:${port}`;
    this.authorization = `Basic ${Buffer.from(`office-partner:${password}`).toString("base64")}`;
    const child = (this.options.spawn ?? spawn)(this.execution.command, ["serve", "--port", String(port), "--hostname", "127.0.0.1"], {
      windowsHide: true, stdio: ["ignore", "pipe", "pipe"], cwd: os.tmpdir(),
      env: { ...this.execution.env, NO_COLOR: "1", FORCE_COLOR: "0", OPENCODE_SERVER_USERNAME: "office-partner", OPENCODE_SERVER_PASSWORD: password },
    });
    this.child = child;
    child.stdout?.resume();
    child.stderr?.resume();
    let failed: string | null = null;
    child.once("error", () => { failed = "Could not run the OpenCode CLI. Install or update it, then try again."; });
    child.once("exit", () => { failed ??= "OpenCode stopped before it was ready."; });
    const deadline = Date.now() + (this.options.startupMs ?? 45_000);
    while (!this.isClosed) {
      if (failed) throw new OpenCodeAuthError(failed);
      if (Date.now() > deadline) throw new OpenCodeAuthError("OpenCode did not start in time. Try again.");
      try { await this.request("GET", "/provider/auth"); return; } catch { await delay(400); }
    }
    throw new OpenCodeAuthError("Sign-in was cancelled.");
  }

  async catalog(): Promise<LoginAccount[]> {
    await this.start();
    const [auth, list] = await Promise.all([
      this.request<Record<string, Record<string, unknown>[]>>("GET", "/provider/auth"),
      this.request<{ all?: { id?: unknown; name?: unknown }[]; connected?: unknown }>("GET", "/provider"),
    ]);
    const connected = new Set(Array.isArray(list.connected) ? list.connected.filter((id): id is string => typeof id === "string") : []);
    const accounts: LoginAccount[] = [];
    const seen = new Set<string>();
    for (const item of list.all ?? []) {
      const id = str(item.id, 80);
      if (!id || seen.has(id)) continue;
      seen.add(id);
      const methods: LoginMethod[] = Array.isArray(auth[id])
        ? auth[id]!.flatMap((method, index): LoginMethod[] => {
          const label = str(method?.label);
          return label && (method.type === "oauth" || method.type === "api") ? [{ index, type: method.type, label, prompts: prompts(method.prompts) }] : [];
        })
        : [];
      // An account with no special sign-in just takes an API key.
      accounts.push({ id, name: str(item.name, 120) ?? id, connected: connected.has(id), methods: methods.length ? methods : [{ index: 0, type: "api", label: "API key", prompts: [] }] });
    }
    const rank = (account: LoginAccount) => { const at = POPULAR.indexOf(account.id); return at < 0 ? POPULAR.length : at; };
    const oauth = (account: LoginAccount) => Number(account.methods.some(method => method.type === "oauth"));
    return accounts.sort((a, b) => rank(a) - rank(b) || oauth(b) - oauth(a) || a.name.localeCompare(b.name));
  }

  async authorize(accountId: string, method: number, inputs: Record<string, string>): Promise<OpenCodeAuthorization> {
    await this.start();
    const result = await this.request<Partial<OpenCodeAuthorization>>("POST", `/provider/${encodeURIComponent(accountId)}/oauth/authorize`,
      { method, ...(Object.keys(inputs).length ? { inputs } : {}) });
    if (typeof result.url !== "string" || (result.method !== "auto" && result.method !== "code")) throw new OpenCodeAuthError("OpenCode did not return a sign-in page.");
    return { url: result.url, method: result.method, instructions: typeof result.instructions === "string" ? result.instructions : "" };
  }

  async complete(accountId: string, method: number, code?: string): Promise<void> {
    await this.start();
    await this.request("POST", `/provider/${encodeURIComponent(accountId)}/oauth/callback`, { method, ...(code === undefined ? {} : { code }) });
  }

  async setKey(accountId: string, key: string, metadata?: Record<string, string>): Promise<void> {
    await this.start();
    await this.request("PUT", `/auth/${encodeURIComponent(accountId)}`, { type: "api", key, ...(metadata && Object.keys(metadata).length ? { metadata } : {}) });
  }

  close(): void {
    this.isClosed = true;
    if (this.idle) clearTimeout(this.idle);
    this.abort.abort();
    killProcessTree(this.child);
    this.child = null;
  }
}

/** Checks what the person picked before anything is started. */
function parseChoice(choice: LoginChoice): { accountId: string; method: number; inputs: Record<string, string>; key?: string } {
  if (!SAFE_ID.test(choice.accountId) || !Number.isInteger(choice.method) || choice.method < 0 || choice.method > 50) throw new ProviderLoginError("Choose an account to sign in with.");
  const inputs = Object.fromEntries(Object.entries(choice.inputs ?? {}).filter(([key, value]) => SAFE_ID.test(key) && typeof value === "string" && value.length <= 400));
  const key = typeof choice.key === "string" ? choice.key.trim() : undefined;
  if (key !== undefined && (key.length < 1 || key.length > 4096 || /[\u0000-\u001f\u007f]/.test(key))) throw new ProviderLoginError("That does not look like an API key. Copy it again.");
  return { accountId: choice.accountId, method: choice.method, inputs, ...(key === undefined ? {} : { key }) };
}

/** Signs in to the OpenCode account the person chose: a page to open (with a code when it has one), or an API key handed straight to OpenCode. */
export class OpenCodeLoginAttempt extends LoginAttempt {
  private submit: ((code: string) => void) | null = null;

  private constructor(private readonly service: OpenCodeAuthApi, choice: ReturnType<typeof parseChoice>) {
    super("opencode", choice.key === undefined ? "Getting a sign-in link…" : "Saving the key…");
    void this.run(choice);
  }

  static start(service: OpenCodeAuthApi, choice: LoginChoice): OpenCodeLoginAttempt {
    return new OpenCodeLoginAttempt(service, parseChoice(choice));
  }

  private fail(error: unknown, fallback: string): void {
    this.finish("error", error instanceof OpenCodeAuthError ? error.message : fallback);
  }

  private async run({ accountId, method, inputs, key }: ReturnType<typeof parseChoice>): Promise<void> {
    try {
      if (key !== undefined) {
        await this.service.setKey(accountId, key, Object.keys(inputs).length ? inputs : undefined);
        return this.finish("success", "Signed in to OpenCode.");
      }
      const auth = await this.service.authorize(accountId, method, inputs);
      if (!this.active) return;
      const code = DEVICE_CODE.exec(auth.instructions)?.[1];
      const pasteBack = auth.method === "code";
      this.waiting({
        userCode: code?.toUpperCase() ?? null, verificationUri: auth.url, needsCode: pasteBack,
        message: pasteBack ? "Sign in on the page, then paste the code it shows below."
          : code ? "Sign in on the page and enter the code. Office Partner will update when you finish."
          : auth.instructions || "Sign in on the page. Office Partner will update when you finish.",
      });
      if (pasteBack) {
        this.submit = value => {
          this.submitted = true;
          void this.service.complete(accountId, method, value).then(() => this.finish("success", "Signed in to OpenCode."),
            error => this.fail(error, "That code was not accepted. It may be incomplete or already used. Start again to get a new code."));
        };
      } else {
        await this.service.complete(accountId, method);
        this.finish("success", "Signed in to OpenCode.");
      }
    } catch (error) {
      this.fail(error, "Sign-in did not finish. Try again.");
    }
  }

  submitCode(code: string): void {
    if (!this.submit || this.current.state !== "waiting" || this.submitted) throw new ProviderLoginError("There is no sign-in waiting for a code.");
    const value = code.trim();
    if (!/^[\w#.~%+/=-]{4,512}$/.test(value)) throw new ProviderLoginError("That does not look like a sign-in code. Copy it again from the page.");
    this.current = { ...this.current, needsCode: false, message: "Checking the code…" };
    this.submit(value);
  }

  protected release(): void {
    this.submit = null;
    this.service.close();
  }
}
