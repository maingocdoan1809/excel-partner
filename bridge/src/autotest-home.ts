import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { AutotestSettingsSchema, type AutotestSettings } from "../../protocol/src/autotest.js";
import { AutotestRuns } from "./autotest-runs.js";

export class AutotestHomeError extends Error {}

/** What a run does when something goes wrong; written on save and added (never overriding a choice) when the kit is refreshed. */
const RUN_DEFAULTS = { stop_on_problem: true, case_retries: 1, max_attempts: 2 };
export interface KitSync { updated: string[]; backup: string | null; skipped: string }

/** Field -> what the person sees on the screen, and how to fix it. Zod's own messages ("Invalid", "Number must be...") mean nothing there. */
const FIELD_HELP: Record<string, string> = {
  project: "Project name: enter a short name, for example the project's name.",
  url: "System address: enter the address you open in the browser, for example http://192.0.2.10 or https://staging.example.com.",
  usernameEnv: "“Variable with the user name” wants the NAME of a variable on this computer (for example QA_USER), not the user name itself. Easier: choose “Use a saved sign-in” above and leave this alone.",
  passwordEnv: "“Variable with the password” wants the NAME of a variable (for example QA_PASSWORD), never the password. Do not type a password here. Easier: choose “Use a saved sign-in”.",
  loginUrl: "Sign-in page: the path of the login page, for example /login.",
  testPrefix: "Prefix of test records: 2 to 30 letters, digits, - or _ with no spaces, for example ZZ_AUTOTEST_.",
  maxRepeat: "Stop a run when the same error happens: enter a number from 2 to 10.",
  signIn: "Saved sign-in: choose one from the list.",
  authType: "Sign-in: choose how the kit signs in.",
};
const explain = (issues: Array<{ path: Array<string | number> }>) =>
  [...new Set(issues.map(issue => FIELD_HELP[String(issue.path[0] ?? "")] ?? `${issue.path.join(".") || "Settings"}: not valid.`))].join(" ");

/** The saved sign-ins of the Testing screen (cookies, kept by the add-in). The kit only ever reads the state file. */
export interface SignInSource { list(): Array<{ name: string; url: string; savedAt: string }>; statePath(name: string): string | null }

type Json = Record<string, unknown>;
const asObject = (value: unknown): Json => (value && typeof value === "object" && !Array.isArray(value) ? value as Json : {});
const text = (value: unknown, fallback = "") => typeof value === "string" ? value : fallback;

const KNOWLEDGE_HEADER = `# Knowledge index

One line per fact, so \`grep\` finds it fast. Format:

    - [screen:<screen>] [kind:<kind>] [tag:<word>] <one-line fact> -> <file or ->

kind = login | navigation | locator | data | error | fix | quirk | rule
Search before acting:   grep -i "screen:orders" knowledge/index.md        grep -i "kind:error" knowledge/index.md | grep -i timeout
Add after learning:      one new line at the END of this file (newest last). Never delete a line that is still true; fix a wrong one in place.
Never write a password, token or cookie here.

`;

export interface KnowledgeLine { screen: string; kind: string; tags: string[]; fact: string; ref: string; line: number }

/** `- [screen:orders] [kind:error] [tag:timeout] fact -> ref` lines of knowledge/index.md. */
export function parseKnowledge(content: string): KnowledgeLine[] {
  const lines: KnowledgeLine[] = [];
  content.split("\n").forEach((raw, index) => {
    const match = /^- ((?:\[[a-z]+:[^\]]+\]\s*)+)(.*)$/.exec(raw.trim());
    if (!match) return;
    const labels = [...match[1]!.matchAll(/\[([a-z]+):([^\]]+)\]/g)].map(item => [item[1]!, item[2]!] as const);
    const [fact, ref = "-"] = match[2]!.split(/\s+->\s+/);
    lines.push({
      screen: labels.find(([key]) => key === "screen")?.[1] ?? "",
      kind: labels.find(([key]) => key === "kind")?.[1] ?? "",
      tags: labels.filter(([key]) => key === "tag").map(([, value]) => value),
      fact: (fact ?? "").trim(), ref: ref.trim(), line: index + 1,
    });
  });
  return lines;
}

/**
 * The single folder that holds everything the testing workflow needs: the kit's configuration (`qa.config.json`, hooks),
 * the kit code, the testcase snapshots, run results and what the assistant has learned. The user picks it once.
 */
export class AutotestHome {
  /** `pointer` is a small JSON file in the data folder that remembers a folder the user chose. */
  constructor(private readonly pointer: string, private readonly defaultDir: string, private readonly kitRoot: string, private readonly skillsRoot: string | null = null, private readonly signIns: SignInSource | null = null) {}

  get dir(): string {
    try {
      const chosen = text(asObject(JSON.parse(fs.readFileSync(this.pointer, "utf8"))).home);
      if (chosen && path.isAbsolute(chosen)) return chosen;
    } catch { /* nothing chosen yet */ }
    return this.defaultDir;
  }

  get configFile(): string { return path.join(this.dir, "qa.config.json"); }
  get outDir(): string { return path.join(this.dir, "test_execution_output"); }
  get testcasesDir(): string { return path.join(this.dir, "testcases"); }
  get knowledgeFile(): string { return path.join(this.dir, "knowledge", "index.md"); }
  get kitVersion(): string { try { return fs.readFileSync(path.join(this.kitRoot, "VERSION"), "utf8").trim(); } catch { return ""; } }
  get kitInstalled(): boolean { return fs.existsSync(path.join(this.dir, "qakit", "__init__.py")); }
  /** Follows the folder the user chose, even when it changes while the bridge runs. */
  runs(): AutotestRuns { return new AutotestRuns(() => this.outDir); }

  /** Where the copy of a file that is about to be replaced goes (set only while the kit is being refreshed on its own). */
  private backupTo: string | null = null;

  private backUp(absTarget: string, label: string): void {
    if (!this.backupTo || !fs.existsSync(absTarget)) return;
    const to = path.join(this.dir, "qakit_backup", this.backupTo, label);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(absTarget, to);
  }

  /** A fingerprint of the kit code and skills this installation carries. */
  private stampCache = "";

  private bundledStamp(): string {
    if (this.stampCache) return this.stampCache;
    const hash = createHash("sha1");
    const walk = (dir: string, rel: string) => {
      for (const entry of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
        if (entry.name === "__pycache__") continue;
        const full = path.join(dir, entry.name), label = `${rel}/${entry.name}`;
        if (entry.isDirectory()) walk(full, label); else { hash.update(label); hash.update(fs.readFileSync(full)); }
      }
    };
    for (const [sub, label] of [["kit/qakit", "qakit"], ["skills", "skills"]] as const) if (fs.existsSync(path.join(this.kitRoot, sub))) walk(path.join(this.kitRoot, sub), label);
    return (this.stampCache = hash.digest("hex").slice(0, 16));
  }

  /**
   * Keep the kit and the skills on this computer the same as the ones this release carries, without anybody pressing Save: the assistant runs
   * whatever is in the folder, so an old copy means old rules. Only for a folder that was set up already; never while a run is going; the copy of
   * every file it replaces goes to qakit_backup/<time>/. The person's own files (hooks, pages, suites, knowledge, config choices) are not touched.
   */
  syncKit(): KitSync {
    if (!this.kitInstalled) return { updated: [], backup: null, skipped: "not set up yet" };
    const stampFile = path.join(this.dir, ".kit-stamp");
    const want = this.bundledStamp();
    try { if (fs.readFileSync(stampFile, "utf8").trim() === want) return { updated: [], backup: null, skipped: "already current" }; } catch { /* never synced */ }
    if (this.runs().list(1)[0]?.phase === "running") return { updated: [], backup: null, skipped: "a run is in progress" };
    const stamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
    let updated: string[] = [];
    this.backupTo = stamp;
    try { updated = [...this.installKit(), ...this.installSkills()]; } finally { this.backupTo = null; }
    const config = this.readConfig();
    if (fs.existsSync(this.configFile)) {
      const run = { ...asObject(config.run) };
      let changed = false;
      for (const [key, value] of Object.entries(RUN_DEFAULTS)) if (run[key] === undefined) { run[key] = value; changed = true; }
      if (changed) { config.run = run; fs.writeFileSync(this.configFile, JSON.stringify(config, null, 1) + "\n", "utf8"); }
    }
    fs.writeFileSync(stampFile, want, "utf8");
    const backup = fs.existsSync(path.join(this.dir, "qakit_backup", stamp)) ? `qakit_backup/${stamp}` : null;
    return { updated, backup, skipped: "" };
  }

  /** Choose another folder. It must be absolute, not a drive or filesystem root, and either new/empty or already an autotest home. */
  choose(dir: string): void {
    const resolved = path.resolve(dir);
    if (!path.isAbsolute(dir) || path.parse(resolved).root === resolved) throw new AutotestHomeError("Choose a normal folder, not a drive root.");
    if (fs.existsSync(resolved)) {
      if (!fs.statSync(resolved).isDirectory()) throw new AutotestHomeError("That path is a file, not a folder.");
      if (fs.readdirSync(resolved).length && !fs.existsSync(path.join(resolved, "qa.config.json")) && !fs.existsSync(path.join(resolved, "qakit"))) {
        throw new AutotestHomeError("That folder is not empty and is not an autotest folder. Choose an empty or new folder.");
      }
    }
    fs.mkdirSync(path.dirname(this.pointer), { recursive: true });
    fs.writeFileSync(this.pointer, JSON.stringify({ home: resolved }, null, 1), "utf8");
  }

  private readConfig(): Json {
    try { return asObject(JSON.parse(fs.readFileSync(this.configFile, "utf8"))); } catch { /* none or unreadable: start from the example */ }
    try { return asObject(JSON.parse(fs.readFileSync(path.join(this.kitRoot, "kit", "qa.config.example.json"), "utf8"))); } catch { return {}; }
  }

  signInList(): Array<{ name: string; url: string; savedAt: string }> { return this.signIns?.list() ?? []; }

  /** The screen's fields as stored in qa.config.json, or null while nothing has been saved. */
  settings(): AutotestSettings | null {
    if (!fs.existsSync(this.configFile)) return null;
    const config = this.readConfig();
    const app = asObject(config.app), auth = asObject(config.auth), data = asObject(config.data), language = asObject(config.language), source = asObject(config.source), run = asObject(config.run);
    const parsed = AutotestSettingsSchema.safeParse({
      project: text(config.project), url: text(app.url), testAccount: text(app.test_account), authType: text(auth.type, "manual"),
      loginUrl: text(auth.login_url, "/login"), usernameEnv: text(auth.username_env, "QA_USER"), passwordEnv: text(auth.password_env, "QA_PASSWORD"),
      testPrefix: text(data.test_prefix, "ZZ_AUTOTEST_"), talkLanguage: text(language.talk, "vi"), blackBox: source.use === false,
      maxRepeat: typeof run.max_repeat === "number" ? run.max_repeat : 3,
      signIn: text(auth.session_name), askBeforeTestData: run.ask_before_test_data === true,
    });
    return parsed.success ? parsed.data : null;
  }

  /** `report.result_values` of the kit config: the words the workbook uses for PASS / FAIL. */
  resultValues(): Record<string, string> {
    const values = asObject(asObject(this.readConfig().report).result_values);
    return Object.fromEntries(Object.entries(values).filter((pair): pair is [string, string] => typeof pair[1] === "string"));
  }

  /** Write the settings into qa.config.json (every other key stays as it was), then install the kit and scaffolding. */
  save(input: unknown): { configFile: string; installed: string[] } {
    const parsed = AutotestSettingsSchema.safeParse(input);
    if (!parsed.success) throw new AutotestHomeError(explain(parsed.error.issues));
    const s = parsed.data;
    const config = this.readConfig();
    const sub = (key: string) => { const value = { ...asObject(config[key]) }; config[key] = value; return value; };
    config.project = s.project;
    Object.assign(sub("language"), { talk: s.talkLanguage });
    Object.assign(sub("app"), { url: s.url, test_account: s.testAccount });
    const auth = sub("auth");
    Object.assign(auth, { type: s.authType, login_url: s.loginUrl, username_env: s.usernameEnv, password_env: s.passwordEnv });
    if (s.authType === "session") {
      const state = s.signIn ? this.signIns?.statePath(s.signIn) : null;
      if (!state) throw new AutotestHomeError(s.signIn ? `The saved sign-in '${s.signIn}' was not found. Sign in again in the Sign-ins section below.` : "Choose a saved sign-in, or sign in first in the Sign-ins section below.");
      Object.assign(auth, { session_name: s.signIn, storage_state: state });
      if (!asObject(auth.logged_in).url_not_contains && !asObject(auth.logged_in).selector) auth.logged_in = { url_not_contains: "/login" };
      Object.assign(sub("browser"), { mode: "launch" });          // a saved sign-in needs the kit's own browser, never an attached one
    }
    Object.assign(sub("data"), { test_prefix: s.testPrefix });
    Object.assign(sub("source"), { use: !s.blackBox });
    // stop_on_problem: repair the case that needs it at once; case_retries: run it once more immediately; max_attempts: tries inside one action (diagnose between)
    Object.assign(sub("run"), { max_repeat: s.maxRepeat, ask_before_test_data: s.askBeforeTestData, ...RUN_DEFAULTS });
    config.testcases = { ...asObject(config.testcases), sources: ["testcases/*.xlsx"] };
    const installed = [...this.installKit(), ...this.installSkills()];
    fs.mkdirSync(this.dir, { recursive: true });
    fs.writeFileSync(this.configFile, JSON.stringify(config, null, 1) + "\n", "utf8");
    fs.writeFileSync(path.join(this.dir, ".kit-stamp"), this.bundledStamp(), "utf8");
    return { configFile: this.configFile, installed };
  }

  /** Copy the kit: framework code (`qakit/`) is refreshed when it differs; project files (hooks, pages, suites, catalog) only when missing. */
  installKit(): string[] {
    const source = path.join(this.kitRoot, "kit");
    if (!fs.existsSync(source)) throw new AutotestHomeError("The testing kit is not part of this installation.");
    const done: string[] = [];
    const walk = (from: string, relative: string) => {
      for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
        const rel = path.join(relative, entry.name);
        if (entry.name === "__pycache__" || entry.name === "qa.config.example.json" && fs.existsSync(path.join(this.dir, entry.name))) continue;
        if (entry.isDirectory()) { walk(path.join(from, entry.name), rel); continue; }
        const target = path.join(this.dir, rel);
        const refresh = rel.split(path.sep)[0] === "qakit";
        if (fs.existsSync(target) && (!refresh || fs.readFileSync(target).equals(fs.readFileSync(path.join(from, entry.name))))) continue;
        this.backUp(target, rel.split(path.sep).join("/"));
        fs.mkdirSync(path.dirname(target), { recursive: true });
        fs.copyFileSync(path.join(from, entry.name), target);
        done.push(rel.split(path.sep).join("/"));
      }
    };
    walk(source, "");
    for (const folder of ["testcases", "knowledge", "test_execution_output"]) fs.mkdirSync(path.join(this.dir, folder), { recursive: true });
    if (!fs.existsSync(this.knowledgeFile)) { fs.writeFileSync(this.knowledgeFile, KNOWLEDGE_HEADER, "utf8"); done.push("knowledge/index.md"); }
    return done;
  }

  /** The kit's skills (system-test, autotest-init) go where the user's own skills live, so the assistant lists and uses them like any skill. */
  installSkills(): string[] {
    const source = path.join(this.kitRoot, "skills");
    if (!this.skillsRoot || !fs.existsSync(source)) return [];
    const done: string[] = [];
    const walk = (from: string, to: string) => {
      fs.mkdirSync(to, { recursive: true });
      for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
        const target = path.join(to, entry.name);
        if (entry.isDirectory()) { walk(path.join(from, entry.name), target); continue; }
        const bytes = fs.readFileSync(path.join(from, entry.name));
        if (fs.existsSync(target) && fs.readFileSync(target).equals(bytes)) continue;
        this.backUp(target, `skills/${path.relative(this.skillsRoot!, target).split(path.sep).join("/")}`);
        fs.writeFileSync(target, bytes);
        done.push(path.relative(this.skillsRoot!, target).split(path.sep).join("/"));
      }
    };
    for (const entry of fs.readdirSync(source, { withFileTypes: true })) if (entry.isDirectory()) walk(path.join(source, entry.name), path.join(this.skillsRoot, entry.name));
    return done;
  }

  knowledge(): KnowledgeLine[] {
    try { return parseKnowledge(fs.readFileSync(this.knowledgeFile, "utf8")); } catch { return []; }
  }

  /**
   * The workbook the user has open, exactly as it is in Excel right now (unsaved edits included). It becomes the kit's testcase
   * source; results go back into the open workbook, never into this copy. `ACTIVE.json` tells the assistant which one is current.
   */
  saveWorkbook(name: string, bytes: Buffer): string {
    const safe = path.basename(name).replace(/[^\w.\- ()À-￿]/g, "_").replace(/\.(xlsx|xlsm)$/i, "") || "workbook";
    if (bytes.length < 100 || bytes.subarray(0, 2).toString("latin1") !== "PK") throw new AutotestHomeError("That is not an Excel workbook.");
    fs.mkdirSync(this.testcasesDir, { recursive: true });
    const file = path.join(this.testcasesDir, `${safe}.xlsx`);
    fs.writeFileSync(file, bytes);
    fs.writeFileSync(path.join(this.testcasesDir, "ACTIVE.json"), JSON.stringify({ file: `${safe}.xlsx`, savedAt: new Date().toISOString() }, null, 1), "utf8");
    return file;
  }

  activeWorkbook(): string | null {
    try { return text(asObject(JSON.parse(fs.readFileSync(path.join(this.testcasesDir, "ACTIVE.json"), "utf8"))).file) || null; } catch { return null; }
  }
}
