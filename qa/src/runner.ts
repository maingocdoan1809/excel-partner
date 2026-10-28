import fs from "node:fs";
import path from "node:path";
import { AUTOMATION_DIR, loadConfig, requireEnvironment, type QaConfig } from "./config.js";
import { checkHarness } from "./harness.js";
import { loadManifest, recordRun, resolveScript, type EffectiveStatus, type Health, type ScriptEntry } from "./manifest.js";
import { playwrightCli, qaEnv, runLogin, runNode } from "./runtime.js";
import { checkSession } from "./session.js";
import { QaError, assertTcId, isInside, readJson } from "./util.js";
import { loadCredentials } from "./vault.js";

export interface RunOptions {
  projectDir: string;
  environment?: string;
  /** Empty = every approved script. */
  tcIds?: string[];
  tag?: string;
  /** Allow scripts that are not approved. They need explicit ids and run visibly unless `headed` is false. */
  draft?: boolean;
  headed?: boolean;
  trace?: boolean;
  signal?: AbortSignal;
  onLine?: (line: string) => void;
  /** Reserve the id up front (see allocateRunId) so callers can report it before the run finishes. */
  runId?: string;
}

export interface PageState {
  url?: string; title?: string; busy: boolean;
  dialogs: Array<{ title: string; text: string; buttons: string[] }>;
  invalid: string[]; notices: string[];
}

export interface RunTestResult {
  tcId: string; title: string; status: "passed" | "failed" | "flaky" | "skipped" | "timedOut";
  durationMs: number; retries: number; error?: string; evidence: string[];
  /** What the page looked like when the test failed (from the harness). Text comes from the system under test. */
  pageState?: PageState; failedAt?: { file: string; line: number; column: number }; failedStep?: string; snippet?: string;
  role: string; req?: string; scriptStatus: EffectiveStatus; health: Health;
}

export interface RunReport {
  runId: string; environment: string; baseUrl: string; startedAt: string; finishedAt: string;
  draftMode: boolean;
  summary: { total: number; passed: number; failed: number; flaky: number; skipped: number };
  tests: RunTestResult[];
  /** Requested or selected scripts that were not run, with the reason. */
  notRun: Array<{ tcId: string; reason: string }>;
  fatal?: string;
  resultsFile: string;
}

const RUN_ID_RE = /^\d{8}-\d{6}(-\d+)?$/;
export const resultsDir = (projectDir: string): string => path.join(projectDir, "results");

/** Reserves a fresh run folder name; the folder is created so concurrent callers cannot get the same id. */
export function allocateRunId(projectDir: string): string {
  const id = newRunId(projectDir);
  fs.mkdirSync(path.join(resultsDir(projectDir), id), { recursive: true });
  return id;
}

function newRunId(projectDir: string): string {
  const now = new Date();
  const pad = (value: number) => String(value).padStart(2, "0");
  const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  let id = stamp;
  for (let n = 2; fs.existsSync(path.join(resultsDir(projectDir), id)); n++) id = `${stamp}-${n}`;
  return id;
}

const healthOf = (status: RunTestResult["status"]): Health => status === "passed" ? "ok" : status === "flaky" ? "flaky" : status === "skipped" ? "unknown" : "failing";
const escapeRegex = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

interface Selected { tcId: string; entry: ScriptEntry; status: EffectiveStatus }

export function selectScripts(projectDir: string, config: QaConfig, options: Pick<RunOptions, "tcIds" | "tag" | "draft">): { selected: Selected[]; notRun: RunReport["notRun"] } {
  const manifest = loadManifest(projectDir);
  const requested = (options.tcIds ?? []).map(assertTcId);
  if (options.draft && requested.length === 0) throw new QaError("Draft scripts only run when you name them (for example: qa run TC_002 --draft).", "invalid_arguments");
  const candidates = requested.length ? requested : Object.keys(manifest.scripts).sort();
  const selected: Selected[] = [], notRun: RunReport["notRun"] = [];
  for (const tcId of candidates) {
    if (!manifest.scripts[tcId]) { notRun.push({ tcId, reason: "no script is registered for this id" }); continue; }
    const resolved = resolveScript(projectDir, config, tcId, manifest);
    if (options.tag && !resolved.entry.tags.includes(options.tag.replace(/^@/, ""))) continue;
    if (resolved.status === "missing") { notRun.push({ tcId, reason: "the script file is missing" }); continue; }
    if (resolved.status === "invalid") { notRun.push({ tcId, reason: `fails validation: ${resolved.validation!.violations[0]?.message}` }); continue; }
    if (resolved.status !== "approved" && !options.draft) {
      // Silent when sweeping "all approved"; explicit when the person asked for this id.
      if (requested.length) notRun.push({ tcId, reason: resolved.status === "modified" ? "changed since it was approved; review it and run: qa approve " + tcId : "not approved yet; review it, watch it pass with --draft, then run: qa approve " + tcId });
      continue;
    }
    selected.push({ tcId, entry: resolved.entry, status: resolved.status });
  }
  return { selected, notRun };
}

/** Make sure every role has a live session, refreshing from the vault when possible. */
export async function ensureSessions(projectDir: string, config: QaConfig, environmentName: string, roles: string[], options: { signal?: AbortSignal; onLine?: (line: string) => void } = {}): Promise<void> {
  const environment = requireEnvironment(config, environmentName);
  const problems: string[] = [];
  for (const role of roles) {
    const status = await checkSession(projectDir, config, environment.name, role);
    if (status.valid) continue;
    if (status.unreachable) { problems.push(`${role}: ${status.reason}`); continue; }
    const credentials = await loadCredentials(environment.baseUrl, role).catch(() => null);
    if (credentials) {
      options.onLine?.(`Session for '${role}' is not valid (${status.reason}); signing in again from the stored password...`);
      try {
        await runLogin({ projectDir, config, environment: environment.name, profile: role, mode: "headless", credentials, signal: options.signal });
        const again = await checkSession(projectDir, config, environment.name, role);
        if (again.valid) continue;
        problems.push(`${role}: signed in but the session is still not accepted (${again.reason})`);
      } catch (error) { problems.push(`${role}: automatic sign-in failed (${(error as Error).message})`); }
      continue;
    }
    problems.push(`${role}: ${status.reason}. Run: excel-partner qa login ${role}${environment.name === config.defaultEnvironment ? "" : ` --env ${environment.name}`}`);
  }
  if (problems.length) throw new QaError(`Cannot run: ${problems.join("; ")}`, "session_invalid");
}

export function pruneResults(projectDir: string, keepDays: number): number {
  const root = resultsDir(projectDir);
  if (!fs.existsSync(root)) return 0;
  const cutoff = Date.now() - keepDays * 86_400_000;
  let removed = 0;
  for (const name of fs.readdirSync(root)) {
    const dir = path.join(root, name);
    if (!RUN_ID_RE.test(name) || !isInside(root, dir)) continue;
    try { if (fs.statSync(dir).mtimeMs < cutoff) { fs.rmSync(dir, { recursive: true, force: true }); removed++; } } catch { /* ignore */ }
  }
  return removed;
}

export async function runTests(options: RunOptions): Promise<RunReport> {
  const { projectDir } = options;
  const config = loadConfig(projectDir);
  const environment = requireEnvironment(config, options.environment);
  const harness = checkHarness(projectDir);
  if (!harness.ok) throw new QaError(`The test harness was changed or is incomplete (${harness.problems.join("; ")}). Review the change, then run: excel-partner qa upgrade`, "harness_modified");
  const cli = playwrightCli(projectDir);

  const { selected, notRun } = selectScripts(projectDir, config, options);
  const startedAt = new Date().toISOString();
  if (options.runId && (!RUN_ID_RE.test(options.runId) || !fs.existsSync(path.join(resultsDir(projectDir), options.runId)))) throw new QaError("Unknown run id.", "invalid_arguments");
  const runId = options.runId ?? newRunId(projectDir);
  const runDir = path.join(resultsDir(projectDir), runId);
  const resultsFile = path.join(runDir, "results.json");
  const empty = (fatal?: string): RunReport => ({
    runId, environment: environment.name, baseUrl: environment.baseUrl, startedAt, finishedAt: new Date().toISOString(), draftMode: Boolean(options.draft),
    summary: { total: 0, passed: 0, failed: 0, flaky: 0, skipped: 0 }, tests: [], notRun, fatal, resultsFile: path.relative(projectDir, resultsFile).split(path.sep).join("/"),
  });
  if (!selected.length) return empty(notRun.length ? undefined : "No approved scripts to run. Approve scripts with: excel-partner qa approve <TC_ID>");

  await ensureSessions(projectDir, config, environment.name, [...new Set(selected.map(item => item.entry.role))], { signal: options.signal, onLine: options.onLine });
  fs.mkdirSync(runDir, { recursive: true });

  const files = selected.map(item => escapeRegex(path.basename(item.entry.file)));
  const headed = options.headed ?? Boolean(options.draft);
  const result = await runNode({
    cwd: projectDir, signal: options.signal, onLine: options.onLine,
    timeoutMs: config.run.totalTimeoutMinutes * 60_000,
    args: [cli, "test", "--config", `${AUTOMATION_DIR}/playwright.config.js`, ...files],
    env: qaEnv(projectDir, config, environment.name, { QA_RUN_DIR: runDir, QA_HEADED: headed ? "1" : "0", QA_TRACE: options.trace ? "1" : "0" }),
  });

  const raw = readJson<{ startedAt?: string; finishedAt?: string; fatal?: string; tests?: Array<Partial<RunTestResult>> }>(resultsFile);
  const tail = (result.stderr || result.stdout).trim().split("\n").slice(-12).join("\n");
  if (!raw) {
    return empty(result.cancelled ? "The run was cancelled." : result.timedOut ? `The run exceeded ${config.run.totalTimeoutMinutes} minutes and was stopped.` : `Playwright produced no results (exit ${result.code}).\n${tail}`);
  }
  const bySelected = new Map(selected.map(item => [item.tcId, item]));
  const tests: RunTestResult[] = (raw.tests ?? []).filter(test => test.tcId && bySelected.has(test.tcId)).map(test => {
    const item = bySelected.get(test.tcId!)!;
    const status = (test.status ?? "failed") as RunTestResult["status"];
    return { tcId: test.tcId!, title: test.title ?? item.entry.title, status, durationMs: test.durationMs ?? 0, retries: test.retries ?? 0, error: test.error, evidence: test.evidence ?? [],
      pageState: test.pageState, failedAt: test.failedAt, failedStep: test.failedStep, snippet: test.snippet,
      role: item.entry.role, req: item.entry.req, scriptStatus: item.status, health: healthOf(status) };
  });
  for (const item of selected) {
    if (!tests.some(test => test.tcId === item.tcId)) notRun.push({ tcId: item.tcId, reason: "Playwright did not report a result for this script (was the test title changed?)" });
  }
  const count = (status: RunTestResult["status"][]) => tests.filter(test => status.includes(test.status)).length;
  recordRun(projectDir, runId, tests.map(test => ({ tcId: test.tcId, status: test.status, durationMs: test.durationMs, health: test.health })));
  pruneResults(projectDir, config.run.keepDays);
  const report: RunReport = {
    runId, environment: environment.name, baseUrl: environment.baseUrl, startedAt, finishedAt: raw.finishedAt ?? new Date().toISOString(), draftMode: Boolean(options.draft),
    summary: { total: tests.length, passed: count(["passed"]), failed: count(["failed", "timedOut"]), flaky: count(["flaky"]), skipped: count(["skipped"]) },
    tests, notRun, fatal: raw.fatal ?? (result.cancelled ? "The run was cancelled." : result.timedOut ? "The run timed out." : undefined),
    resultsFile: path.relative(projectDir, resultsFile).split(path.sep).join("/"),
  };
  fs.writeFileSync(path.join(runDir, "report.json"), JSON.stringify(report, null, 2));
  return report;
}

export function latestRunId(projectDir: string): string | null {
  const root = resultsDir(projectDir);
  if (!fs.existsSync(root)) return null;
  const ids = fs.readdirSync(root).filter(name => RUN_ID_RE.test(name) && fs.existsSync(path.join(root, name, "report.json"))).sort();
  return ids.at(-1) ?? null;
}

export function readReport(projectDir: string, runId?: string): RunReport {
  const id = runId ?? latestRunId(projectDir);
  if (!id || !RUN_ID_RE.test(id)) throw new QaError(runId ? `'${runId}' is not a run id.` : "No runs yet. Run the tests first.", "no_run");
  const report = readJson<RunReport>(path.join(resultsDir(projectDir), id, "report.json"));
  if (!report) throw new QaError(`Run ${id} has no report.`, "no_run");
  return report;
}
