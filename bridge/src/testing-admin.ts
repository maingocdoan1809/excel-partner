import { execFile } from "node:child_process";
import { normalizeOrigin } from "../../protocol/src/tools-testing.js";
import { loginSession, type LoginOptions } from "./test-login.js";
import { SESSION_NAME_PATTERN, type TestSessionInfo, type TestSessions } from "./test-sessions.js";
import { originOfKey, type KnowledgeEntry, type SiteKnowledge } from "./site-knowledge.js";
import type { TestScriptEntry, TestScripts } from "./test-scripts.js";

export class TestingAdminError extends Error {}

export interface TestingOverview {
  signIns: TestSessionInfo[];
  notes: Array<{ site: string; notes: KnowledgeEntry[] }>;
  scripts: Array<{ site: string; scripts: TestScriptEntry[] }>;
  /** Set while a browser window is open for a sign-in. */
  pendingSignIn: { name: string; url: string } | null;
  folders: { sessions: string; knowledge: string; scripts: string };
}

interface PendingSignIn { name: string; url: string; finish(): void; cancel(): void; done: Promise<{ origins: string[] }>; timer: NodeJS.Timeout }

export type FolderKind = "sessions" | "knowledge" | "scripts";
export const SIGN_IN_TIMEOUT_MS = 10 * 60_000;

/**
 * What the `excel-partner login/sessions/logout/knowledge/scripts` commands do, behind buttons: list, sign in with a
 * visible browser, view and delete. It touches only the three folders next to the configuration; nothing here takes a path.
 */
export class TestingAdmin {
  private pending: PendingSignIn | null = null;

  constructor(
    private readonly sessions: TestSessions,
    private readonly knowledge: SiteKnowledge,
    private readonly scripts: TestScripts,
    private readonly login: (options: LoginOptions) => Promise<{ origins: string[] }> = loginSession,
    private readonly openFolder: (dir: string) => void = dir => { execFile("explorer.exe", [dir], () => undefined); },
    private readonly timeoutMs = SIGN_IN_TIMEOUT_MS,
  ) {}

  overview(): TestingOverview {
    const listed = <T>(dirs: string[], read: (origin: string) => T[]) => dirs.flatMap(site => {
      const origin = originOfKey(site);
      const items = origin ? read(origin) : [];
      return items.length ? [{ site, items }] : [];
    });
    return {
      signIns: this.sessions.list(),
      notes: listed(this.knowledge.sites(), origin => this.knowledge.list(origin)).map(({ site, items }) => ({ site, notes: items })),
      scripts: listed(this.scripts.sites(), origin => this.scripts.list(origin)).map(({ site, items }) => ({ site, scripts: items })),
      pendingSignIn: this.pending ? { name: this.pending.name, url: this.pending.url } : null,
      folders: { sessions: this.sessions.dir, knowledge: this.knowledge.dir, scripts: this.scripts.dir },
    };
  }

  private originOf(site: string): string {
    const origin = originOfKey(site);
    if (!origin) throw new TestingAdminError(`'${site}' is not a known site folder.`);
    return origin;
  }

  viewNote(site: string, name: string): { entry: KnowledgeEntry; content: string } {
    try { return this.knowledge.read(this.originOf(site), name); } catch (error) { throw new TestingAdminError(error instanceof Error ? error.message : String(error)); }
  }

  viewScript(site: string, caseId: string): string {
    try { return this.scripts.read(this.originOf(site), caseId); } catch (error) { throw new TestingAdminError(error instanceof Error ? error.message : String(error)); }
  }

  deleteNote(site: string, name: string): boolean {
    try { return this.knowledge.delete(this.originOf(site), name); } catch (error) { throw new TestingAdminError(error instanceof Error ? error.message : String(error)); }
  }

  deleteScript(site: string, caseId: string): boolean {
    try { return this.scripts.delete(this.originOf(site), caseId); } catch (error) { throw new TestingAdminError(error instanceof Error ? error.message : String(error)); }
  }

  deleteSignIn(name: string): boolean {
    if (!SESSION_NAME_PATTERN.test(name)) throw new TestingAdminError(`'${name}' is not a valid sign-in name.`);
    return this.sessions.remove(name);
  }

  /** Open a visible browser at the address; resolves once the window is open and the person can sign in. */
  async startSignIn(input: { name: string; url: string; extraOrigins?: string[] }): Promise<{ name: string; url: string }> {
    if (this.pending) throw new TestingAdminError(`A browser is already open for the sign-in '${this.pending.name}'. Finish or cancel it first.`);
    if (!SESSION_NAME_PATTERN.test(input.name)) throw new TestingAdminError("The name may use lowercase letters, digits, - and _ (up to 32 characters), for example admin or tester-1.");
    const origin = normalizeOrigin(input.url);
    if (!origin) throw new TestingAdminError("Enter the system's address, for example https://staging.example.com.");
    let finish!: () => void;
    let cancel!: (error: Error) => void;
    const gate = new Promise<void>((resolve, reject) => { finish = resolve; cancel = reject; });
    gate.catch(() => undefined);
    let opened!: () => void;
    const isOpen = new Promise<void>(resolve => { opened = resolve; });
    const done = this.login({ name: input.name, url: origin, extraOrigins: input.extraOrigins, store: this.sessions, waitForPerson: () => gate, out: () => opened() });
    done.catch(() => undefined);
    const pending: PendingSignIn = {
      name: input.name, url: origin, finish, done,
      cancel: () => cancel(new Error("Sign-in cancelled.")),
      timer: setTimeout(() => cancel(new Error("The sign-in window was open for too long and was closed.")), this.timeoutMs),
    };
    this.pending = pending;
    const clear = () => { clearTimeout(pending.timer); if (this.pending === pending) this.pending = null; };
    done.then(clear, clear);
    try { await Promise.race([isOpen, done.then(() => undefined)]); } catch (error) { clear(); throw new TestingAdminError(error instanceof Error ? error.message : String(error)); }
    return { name: input.name, url: origin };
  }

  /** The person says they are signed in: save the session and close the window. */
  async finishSignIn(): Promise<{ name: string; origins: string[] }> {
    const pending = this.pending;
    if (!pending) throw new TestingAdminError("No sign-in window is open. Start one first.");
    pending.finish();
    try { return { name: pending.name, ...(await pending.done) }; } catch (error) { throw new TestingAdminError(error instanceof Error ? error.message : String(error)); }
  }

  cancelSignIn(): boolean {
    if (!this.pending) return false;
    this.pending.cancel();
    return true;
  }

  folder(kind: FolderKind): string {
    const dir = kind === "sessions" ? this.sessions.dir : kind === "knowledge" ? this.knowledge.dir : this.scripts.dir;
    if (process.platform !== "win32") throw new TestingAdminError(`Open this folder yourself: ${dir}`);
    this.openFolder(dir);
    return dir;
  }
}
