import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { readActiveProject, setActiveProject } from "./active.js";
import { CONFIG_FILE, allowedOriginsOf, loadConfig, profileNames, requireEnvironment, requireProfile } from "./config.js";
import { HARNESS_VERSION, checkHarness, installHarness, installSkills, templateWorkbookPath } from "./harness.js";
import { MAX_HEAL_ATTEMPTS, approve, listScripts, loadManifest, readApprovedSnapshot, resetHeal, resolveScript, revoke } from "./manifest.js";
import { lineDiff } from "./diff.js";
import { compareAssertions } from "./validator.js";
import { initProject, installDependencies } from "./scaffold.js";
import { checkSession } from "./session.js";
import { runLogin } from "./runtime.js";
import { pruneResults, runTests } from "./runner.js";
import { QaError, assertTcId, resolveInside } from "./util.js";
import { clearCredentials, dpapiAvailable, envNames, hasCredentials, promptHidden, promptLine, storeCredentials } from "./vault.js";
import { findProject } from "./config.js";

export interface CliContext {
  cwd: string;
  /** The bridge's config.json, where the active project is recorded. */
  configFile: string;
  out: (line: string) => void;
  err: (line: string) => void;
  promptLine: (question: string) => Promise<string>;
  promptHidden: (question: string) => Promise<string>;
  isInteractive: boolean;
}

export const QA_HELP = `Usage: excel-partner qa <command>

Set up
  init [dir]            Create a QA project (--preset odoo|generic, --url <https://host>, --no-install)
  use <dir>             Make this the project the Excel add-in works with
  doctor [--fix]        Check the project, harness, Playwright and sessions
  upgrade               Restore the managed harness files to this version
  skills install       Install/update the /qa-testcases /qa-automate /qa-heal /qa-report skills (--force replaces edited ones)
  template [file]       Write a blank TestPlan.xlsx (requirements, test cases, RTM, dashboard, bugs)

Sign in (you type the password; the AI never sees it)
  login <profile>       Open a browser, sign in, save the session (--env <name>)
  secret set <profile>  Store a password, encrypted for your Windows user, to refresh sessions unattended
  secret clear <profile>
  status                Show whether each saved session still works

Scripts
  list                  Scripts and their status: draft | approved | modified | invalid
  validate [TC_ID...]   Run the safety checks
  approve <TC_ID...>    Approve after you have read the script. A script approved before shows what changed since;
                        if checks (assertions) were dropped or changed, --yes is refused unless you add --accept-assertion-changes
  revoke <TC_ID...>     Withdraw approval
  heal-reset <TC_ID...> Allow the assistant more repair attempts for a script (it gets ${MAX_HEAL_ATTEMPTS} per script until it passes or you approve)
  run [TC_ID...]        Run approved scripts (--tag <t>, --env <name>, --trace)
                        Draft scripts: run TC_ID --draft [--headed|--no-headed]
  clean [--days N]      Delete old results

Options: --project <dir> to act on a project other than the current folder.`;

interface Parsed { positional: string[]; flags: Map<string, string | true> }
const VALUE_FLAGS = new Set(["preset", "url", "env", "project", "tag", "days", "environment"]);

function parse(argv: string[]): Parsed {
  const positional: string[] = [], flags = new Map<string, string | true>();
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]!;
    if (arg.startsWith("--")) {
      const [name, inline] = arg.slice(2).split("=", 2) as [string, string | undefined];
      if (inline !== undefined) flags.set(name, inline);
      else if (VALUE_FLAGS.has(name)) { const next = argv[++i]; if (next === undefined) throw new QaError(`--${name} needs a value.`, "invalid_arguments"); flags.set(name, next); }
      else flags.set(name, true);
    } else positional.push(arg);
  }
  return { positional, flags };
}

function projectFor(context: CliContext, flags: Map<string, string | true>): string {
  const explicit = flags.get("project");
  if (typeof explicit === "string") {
    const dir = path.resolve(context.cwd, explicit);
    if (!fs.existsSync(path.join(dir, CONFIG_FILE))) throw new QaError(`${dir} has no ${CONFIG_FILE}.`, "no_project");
    return dir;
  }
  const found = findProject(context.cwd) ?? readActiveProject(context.configFile);
  if (!found || !fs.existsSync(path.join(found, CONFIG_FILE))) throw new QaError("No QA project here. Run: excel-partner qa init <folder>   (or: qa use <folder>)", "no_project");
  return found;
}

const pad = (text: string, width: number) => text.length >= width ? text : text + " ".repeat(width - text.length);
const envFlag = (flags: Map<string, string | true>): string | undefined => typeof flags.get("env") === "string" ? flags.get("env") as string : undefined;

export async function runQaCli(argv: string[], context: CliContext): Promise<number> {
  try {
    const { positional, flags } = parse(argv);
    const [command = "help", ...rest] = positional;
    switch (command) {
      case "help": context.out(QA_HELP); return 0;
      case "init": return await init(context, rest, flags);
      case "use": {
        const dir = path.resolve(context.cwd, rest[0] ?? ".");
        setActiveProject(context.configFile, dir);
        context.out(`The Excel add-in will use ${dir}. Restart the bridge to apply: excel-partner restart`);
        return 0;
      }
      case "skills": {
        if (rest[0] !== "install") throw new QaError("Usage: qa skills install [--force]", "invalid_arguments");
        for (const item of installSkills(flags.has("force"))) context.out(`/${item.name}: ${item.action}${item.action === "kept" ? " (you edited it; use --force to replace)" : ""}`);
        return 0;
      }
      case "template": {
        const target = path.resolve(context.cwd, rest[0] ?? "TestPlan.xlsx");
        if (fs.existsSync(target)) throw new QaError(`${target} already exists.`, "exists");
        fs.copyFileSync(templateWorkbookPath(), target);
        context.out(`Created ${target}`);
        return 0;
      }
      case "login": return await login(context, rest, flags);
      case "secret": return await secret(context, rest, flags);
      case "status": return await status(context, flags);
      case "list": return list(context, flags);
      case "validate": return validate(context, rest, flags);
      case "approve": return await approveCommand(context, rest, flags);
      case "heal-reset": {
        const dir = projectFor(context, flags);
        if (!rest.length) throw new QaError("Name the test case ids.", "invalid_arguments");
        for (const id of rest) { resetHeal(dir, assertTcId(id)); context.out(`${id}: the assistant may repair it again (up to ${MAX_HEAL_ATTEMPTS} more attempts).`); }
        return 0;
      }
      case "revoke": {
        const dir = projectFor(context, flags);
        if (!rest.length) throw new QaError("Name the test case ids to revoke.", "invalid_arguments");
        for (const id of rest) { revoke(dir, assertTcId(id)); context.out(`${id}: approval withdrawn (now a draft).`); }
        return 0;
      }
      case "run": return await run(context, rest, flags);
      case "doctor": return await doctor(context, flags);
      case "upgrade": {
        const dir = projectFor(context, flags);
        installHarness(dir);
        context.out(`Harness restored to version ${HARNESS_VERSION}. Approvals are unaffected.`);
        for (const item of installSkills()) if (item.action !== "unchanged") context.out(`Skill /${item.name}: ${item.action}${item.action === "kept" ? " (you edited it; use 'skills install --force' to replace)" : ""}`);
        return 0;
      }
      case "clean": {
        const dir = projectFor(context, flags);
        const days = typeof flags.get("days") === "string" ? Number(flags.get("days")) : loadConfig(dir).run.keepDays;
        if (!Number.isFinite(days) || days < 0) throw new QaError("--days must be a number.", "invalid_arguments");
        context.out(`Removed ${pruneResults(dir, days)} old run(s).`);
        return 0;
      }
      default: context.err(`Unknown qa command: ${command}\n\n${QA_HELP}`); return 2;
    }
  } catch (error) {
    if (error instanceof QaError) { context.err(error.message); return error.code === "cancelled" ? 130 : 1; }
    throw error;
  }
}

async function init(context: CliContext, rest: string[], flags: Map<string, string | true>): Promise<number> {
  const dir = path.resolve(context.cwd, rest[0] ?? "qa-automation");
  const preset = (flags.get("preset") ?? "odoo") as string;
  if (preset !== "odoo" && preset !== "generic") throw new QaError("--preset must be odoo or generic.", "invalid_arguments");
  let baseUrl = typeof flags.get("url") === "string" ? flags.get("url") as string : "";
  if (!baseUrl && context.isInteractive) baseUrl = (await context.promptLine("Base URL of the system under test (staging, e.g. https://staging.example.com): ")).trim();
  if (!baseUrl) throw new QaError("Pass --url <https://host> (the staging system, never production).", "invalid_arguments");
  let origin: string;
  try { origin = new URL(baseUrl).origin; } catch { throw new QaError(`'${baseUrl}' is not a URL.`, "invalid_arguments"); }
  initProject({ dir, preset, baseUrl: origin, install: !flags.has("no-install"), skills: !flags.has("no-skills"), log: context.out });
  setActiveProject(context.configFile, dir);
  context.out("");
  context.out(`QA project ready: ${dir}`);
  context.out(`Allowed origin (the browser can reach nothing else): ${origin}`);
  context.out("Next:");
  context.out(`  cd "${dir}"`);
  context.out("  excel-partner qa login admin     # you sign in once; the session is saved");
  context.out("  excel-partner qa run TC_000 --draft   # watch the example script run");
  context.out(`  Open ${path.join(dir, "TestPlan.xlsx")} in Excel and use the /qa-testcases skill.`);
  context.out("Then restart the bridge (excel-partner restart), start a new chat, and ask the assistant to write test cases from your workbook.");
  return 0;
}

async function login(context: CliContext, rest: string[], flags: Map<string, string | true>): Promise<number> {
  const dir = projectFor(context, flags);
  const config = loadConfig(dir);
  const profile = rest[0] ?? (profileNames(config).length === 1 ? profileNames(config)[0] : undefined);
  if (!profile) throw new QaError(`Say which profile: ${profileNames(config).join(", ")}`, "invalid_arguments");
  requireProfile(config, profile);
  const environment = requireEnvironment(config, envFlag(flags));
  context.out(`Signing in as '${profile}' on ${environment.baseUrl}`);
  await runLogin({ projectDir: dir, config, environment: environment.name, profile, mode: "interactive" });
  const state = await checkSession(dir, config, environment.name, profile);
  context.out(state.valid ? `Session for '${profile}' saved and accepted by the server.` : `Session saved, but the server does not accept it: ${state.reason}`);
  return state.valid ? 0 : 1;
}

async function secret(context: CliContext, rest: string[], flags: Map<string, string | true>): Promise<number> {
  const [action, profile] = rest;
  if ((action !== "set" && action !== "clear") || !profile) throw new QaError("Usage: qa secret set|clear <profile>", "invalid_arguments");
  const dir = projectFor(context, flags);
  const config = loadConfig(dir);
  const settings = requireProfile(config, profile);
  const environment = requireEnvironment(config, envFlag(flags));
  if (action === "clear") {
    context.out(clearCredentials(environment.baseUrl, profile) ? `Stored password for '${profile}' removed.` : `No stored password for '${profile}'.`);
    return 0;
  }
  if (!dpapiAvailable()) {
    const names = envNames(profile);
    throw new QaError(`Encrypted storage needs Windows. Set the environment variables ${names.user} and ${names.password} instead.`, "unsupported");
  }
  const user = (settings.username || (await context.promptLine(`User name for '${profile}' on ${environment.baseUrl}: `))).trim();
  const password = await context.promptHidden(`Password for ${user} (hidden, stored encrypted for this Windows user only): `);
  await storeCredentials(environment.baseUrl, profile, { user, password });
  context.out(`Stored. Expired sessions for '${profile}' will now be refreshed automatically.`);
  return 0;
}

async function status(context: CliContext, flags: Map<string, string | true>): Promise<number> {
  const dir = projectFor(context, flags);
  const config = loadConfig(dir);
  const environment = requireEnvironment(config, envFlag(flags));
  context.out(`Environment ${environment.name}: ${environment.baseUrl}`);
  let bad = 0;
  for (const profile of profileNames(config)) {
    const state = await checkSession(dir, config, environment.name, profile);
    if (!state.valid) bad++;
    const stored = hasCredentials(environment.baseUrl, profile) ? "password stored" : "no stored password";
    context.out(`  ${pad(profile, 14)} ${state.valid ? "valid  " : "INVALID"} ${state.valid ? "" : `(${state.reason}) `}[${stored}]`);
  }
  return bad ? 1 : 0;
}

function list(context: CliContext, flags: Map<string, string | true>): number {
  const dir = projectFor(context, flags);
  const scripts = listScripts(dir, loadConfig(dir));
  if (!scripts.length) { context.out("No scripts yet."); return 0; }
  context.out(`${pad("TC", 14)}${pad("STATUS", 10)}${pad("LAST", 9)}${pad("ROLE", 10)}TITLE`);
  for (const script of scripts) {
    context.out(`${pad(script.tcId, 14)}${pad(script.status, 10)}${pad(script.health, 9)}${pad(script.role, 10)}${script.title}${script.healAttempts ? `  [repairs ${script.healAttempts}/${MAX_HEAL_ATTEMPTS}]` : ""}${script.problem ? `  !! ${script.problem}` : ""}`);
  }
  return 0;
}

function validate(context: CliContext, rest: string[], flags: Map<string, string | true>): number {
  const dir = projectFor(context, flags);
  const config = loadConfig(dir);
  const manifest = loadManifest(dir);
  const ids = rest.length ? rest.map(assertTcId) : Object.keys(manifest.scripts).sort();
  let failures = 0;
  for (const id of ids) {
    const resolved = resolveScript(dir, config, id, manifest);
    if (resolved.status === "invalid") {
      failures++;
      context.out(`${id}: FAIL`);
      for (const violation of resolved.validation!.violations) context.out(`   ${violation.line ? `line ${violation.line}: ` : ""}${violation.message}`);
    } else if (resolved.status === "missing") { failures++; context.out(`${id}: FAIL script file is missing`); }
    else context.out(`${id}: ok (${resolved.status})`);
  }
  return failures ? 1 : 0;
}

async function approveCommand(context: CliContext, rest: string[], flags: Map<string, string | true>): Promise<number> {
  const dir = projectFor(context, flags);
  const config = loadConfig(dir);
  if (!rest.length) throw new QaError("Name the test case ids to approve.", "invalid_arguments");
  for (const id of rest.map(assertTcId)) {
    const resolved = resolveScript(dir, config, id);
    if (resolved.status === "invalid") throw new QaError(`${id} fails validation: ${resolved.validation!.violations[0]?.message}`, "invalid_script");
    if (resolved.status === "missing") throw new QaError(`${id}: the script file is missing.`, "missing_script");
    if (resolved.status === "approved") { context.out(`${id}: already approved.`); continue; }
    const code = fs.readFileSync(resolveInside(dir, resolved.entry.file), "utf8").replace(/\r\n/g, "\n");
    // A script that was approved before and has been changed since: show what changed, and whether any check was dropped.
    const before = readApprovedSnapshot(dir, id);
    const changed = before !== null && before !== code;
    const comparison = changed ? compareAssertions(before!, code) : null;
    const weakened = Boolean(comparison?.removed.length);
    const acceptWeakened = flags.has("accept-assertion-changes");
    if (weakened && !acceptWeakened && (flags.has("yes") || !context.isInteractive)) {
      throw new QaError(`${id}: checks were removed or changed since the version you approved (${comparison!.removed.slice(0, 3).join("; ")}). Look at the change with 'excel-partner qa approve ${id}' in a terminal, or pass --accept-assertion-changes once you have reviewed it.`, "needs_confirmation");
    }
    if (!flags.has("yes")) {
      if (!context.isInteractive) throw new QaError("Approval needs a person at a terminal (or --yes after you reviewed the script).", "needs_confirmation");
      context.out(`--- ${resolved.entry.file}  (role: ${resolved.entry.role}${resolved.validation!.deps.length ? `, uses ${resolved.validation!.deps.join(", ")}` : ""})`);
      if (changed) {
        const lines = lineDiff(before!, code);
        context.out(`Changes since the version you approved ("-" removed, "+" added):`);
        context.out(lines.slice(0, 200).join("\n"));
        if (lines.length > 200) context.out(`... (${lines.length - 200} more lines; read the file before approving)`);
        if (comparison!.added.length) context.out(`Checks added: ${comparison!.added.join("; ")}`);
      } else {
        context.out(code.split("\n").slice(0, 200).join("\n"));
        if (code.split("\n").length > 200) context.out(`... (${code.split("\n").length - 200} more lines; read the file before approving)`);
      }
      context.out("---");
      if (weakened) context.out(`WARNING: these checks were removed or changed, so the test may now accept something it used to reject:\n  - ${comparison!.removed.join("\n  - ")}`);
      if (resolved.entry.lastRun?.health !== "ok") context.out("Note: this script has not passed a run yet. Try: excel-partner qa run " + id + " --draft");
      if (weakened && !acceptWeakened) {
        const typed = (await context.promptLine(`Type "accept" to approve ${id} with these checks changed (anything else cancels): `)).trim().toLowerCase();
        if (typed !== "accept") { context.out(`${id}: not approved.`); continue; }
      } else {
        const answer = (await context.promptLine(`Approve ${id}? It will then run unattended with the '${resolved.entry.role}' session. [y/N] `)).trim().toLowerCase();
        if (answer !== "y" && answer !== "yes") { context.out(`${id}: not approved.`); continue; }
      }
    }
    approve(dir, config, id, os.userInfo().username);
    context.out(`${id}: approved.`);
  }
  return 0;
}

async function run(context: CliContext, rest: string[], flags: Map<string, string | true>): Promise<number> {
  const dir = projectFor(context, flags);
  const draft = flags.has("draft");
  const headed = flags.has("headed") ? true : flags.has("no-headed") ? false : undefined;
  const report = await runTests({
    projectDir: dir, environment: envFlag(flags), tcIds: rest, draft, headed, trace: flags.has("trace"),
    tag: typeof flags.get("tag") === "string" ? flags.get("tag") as string : undefined, onLine: line => context.out(line),
  });
  context.out("");
  for (const test of report.tests) context.out(`${pad(test.status.toUpperCase(), 9)}${pad(test.tcId, 14)}${(test.durationMs / 1000).toFixed(1)}s${test.error ? `\n          ${test.error.split("\n")[0]}` : ""}`);
  for (const item of report.notRun) context.out(`NOT RUN  ${pad(item.tcId, 14)}${item.reason}`);
  if (report.fatal) context.err(report.fatal);
  const { summary } = report;
  context.out(`Run ${report.runId}: ${summary.passed} passed, ${summary.failed} failed, ${summary.flaky} flaky, ${summary.skipped} skipped.  Results: ${report.resultsFile}`);
  return report.fatal || summary.failed ? 1 : 0;
}

async function doctor(context: CliContext, flags: Map<string, string | true>): Promise<number> {
  const line = (ok: boolean | null, text: string) => context.out(`${ok === true ? "[ ok ]" : ok === false ? "[fail]" : "[warn]"} ${text}`);
  let failed = false;
  const fail = (ok: boolean | null, text: string) => { if (ok === false) failed = true; line(ok, text); };
  fail(Number(process.versions.node.split(".")[0]) >= 20, `Node.js ${process.versions.node}`);
  const dir = projectFor(context, flags);
  fail(true, `Project ${dir}`);
  const config = loadConfig(dir);
  const environment = requireEnvironment(config, envFlag(flags));
  fail(true, `Environment ${environment.name} -> ${environment.baseUrl}`);
  fail(true, `Browser may reach only: ${allowedOriginsOf(config).join(", ")}`);
  if (/production|prod\b|\blive\b/i.test(environment.name + environment.baseUrl)) fail(null, "This looks like a production system. Point tests at staging.");
  const harness = checkHarness(dir);
  fail(harness.ok, harness.ok ? `Harness files intact${harness.outdated ? " (an update is available: qa upgrade)" : ""}` : `Harness problem: ${harness.problems.join("; ")}`);
  const playwright = fs.existsSync(path.join(dir, "node_modules", "@playwright", "test", "cli.js"));
  if (!playwright && flags.has("fix")) { installDependencies(dir, context.out); }
  const installed = fs.existsSync(path.join(dir, "node_modules", "@playwright", "test", "cli.js"));
  fail(installed, installed ? "Playwright installed in the project" : "Playwright is not installed. Run: excel-partner qa doctor --fix");
  const ignore = fs.existsSync(path.join(dir, ".gitignore")) ? fs.readFileSync(path.join(dir, ".gitignore"), "utf8") : "";
  fail(/^auth\/?$/m.test(ignore) && /^results\/?$/m.test(ignore), ".gitignore keeps auth/ and results/ out of git");
  const scripts = listScripts(dir, config);
  fail(!scripts.some(script => script.status === "invalid" || script.status === "missing"), `${scripts.length} script(s): ${scripts.filter(script => script.status === "approved").length} approved, ${scripts.filter(script => script.status === "invalid" || script.status === "missing").length} broken`);
  for (const profile of profileNames(config)) {
    const state = await checkSession(dir, config, environment.name, profile);
    fail(state.valid ? true : null, `Session '${profile}': ${state.valid ? "valid" : `${state.reason}. Run: excel-partner qa login ${profile}`}`);
  }
  return failed ? 1 : 0;
}

export function defaultContext(configFile: string): CliContext {
  return {
    cwd: process.cwd(), configFile, out: line => console.log(line), err: line => console.error(line),
    promptLine: question => promptLine(question), promptHidden: question => promptHidden(question), isInteractive: Boolean(process.stdin.isTTY),
  };
}
