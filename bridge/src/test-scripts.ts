import fs from "node:fs";
import path from "node:path";
import { CASE_ID_PATTERN, normalizeOrigin, type ScriptOutcome } from "../../protocol/src/tools-testing.js";
import { findSecret, siteKey } from "./site-knowledge.js";

export const MAX_SCRIPTS_PER_SITE = 200;

export interface TestScriptEntry { caseId: string; title: string; outcome: ScriptOutcome; tested: string; chars: number }

export class TestScriptError extends Error {}

/** What a regression script may not contain. This is a guard against page text steering the assistant, not a sandbox. */
const FORBIDDEN: Array<[RegExp, string]> = [
  [/\brequire\s*\(/, "require()"],
  [/\bimport\s*\(/, "dynamic import()"],
  [/\bchild_process\b|\bnode:/, "Node modules"],
  [/\beval\s*\(|\bnew\s+Function\b|\bFunction\s*\(/, "eval or Function"],
  [/\bfetch\s*\(|\bXMLHttpRequest\b|\bWebSocket\b|\bsendBeacon\b/, "direct network calls"],
  [/\bprocess\s*\.\s*(?!env\s*\.\s*(?:BASE_URL|STORAGE_STATE)\b)/, "process (only process.env.BASE_URL and process.env.STORAGE_STATE are allowed)"],
  [/\bprocess\s*\[/, "process[...]"],
];

/** Check a script without writing it; save() repeats these checks. */
export function validateScript(origin: string, script: string): void {
  const site = normalizeOrigin(origin);
  if (!site) throw new TestScriptError(`'${origin}' is not a web address (use https://host, with the port if any).`);
  const imports = [...script.matchAll(/^\s*import\s[^;]*?\bfrom\s+['"]([^'"]+)['"]/gm)].map(match => match[1]!);
  if (!imports.includes("@playwright/test")) throw new TestScriptError("The script must start with: import { test, expect } from '@playwright/test';");
  if (imports.some(name => name !== "@playwright/test")) throw new TestScriptError("The script may import only '@playwright/test'.");
  if (/^\s*import\s+['"]/m.test(script)) throw new TestScriptError("The script may import only '@playwright/test'.");
  if (!/\btest\s*\(/.test(script)) throw new TestScriptError("The script needs at least one test('...', async ({ page }) => { ... }).");
  if (!/\bexpect\s*[(.]/.test(script)) throw new TestScriptError("The script never checks anything. Assert the expected result of the test case with expect(...); a script that only clicks proves nothing.");
  for (const [pattern, what] of FORBIDDEN) if (pattern.test(script)) throw new TestScriptError(`The script uses ${what}, which is not allowed in a regression script. Use only Playwright page actions and expect().`);
  for (const match of script.matchAll(/https?:\/\/[^\s'"`)]+/g)) {
    const found = normalizeOrigin(match[0]);
    if (found && found !== site) throw new TestScriptError(`The script opens ${found}, which is not the system under test (${site}). Use relative paths with page.goto('/path'); the base address comes from BASE_URL.`);
  }
  const secret = findSecret(script);
  if (secret) throw new TestScriptError(`The script looks like it contains ${secret}. Never put passwords, tokens or cookies in a script: it runs with the saved session (STORAGE_STATE).`);
}

const oneLine = (text: string) => text.replace(/\s+/g, " ").trim();

/**
 * Playwright scripts for the test cases the assistant has run, kept as plain files on this computer so the person can read,
 * edit, run and delete them. Only cases that passed or that failed because the application is wrong are kept.
 */
export class TestScripts {
  constructor(readonly dir: string) {}

  private siteDir(origin: string): string {
    const key = siteKey(origin);
    if (!key) throw new TestScriptError(`'${origin}' is not a web address (use https://host, with the port if any).`);
    return path.join(this.dir, key);
  }

  file(origin: string, caseId: string): string {
    if (!CASE_ID_PATTERN.test(caseId)) throw new TestScriptError(`'${caseId}' is not a valid test case ID (letters, digits, . _ - up to 64 characters).`);
    return path.join(this.siteDir(origin), `${caseId}.spec.js`);
  }

  list(origin: string): TestScriptEntry[] {
    const dir = this.siteDir(origin);
    let files: string[] = [];
    try { files = fs.readdirSync(dir).filter(file => file.endsWith(".spec.js")); } catch { return []; }
    const entries: TestScriptEntry[] = [];
    for (const file of files) {
      const caseId = file.slice(0, -".spec.js".length);
      if (!CASE_ID_PATTERN.test(caseId)) continue;
      try {
        const text = fs.readFileSync(path.join(dir, file), "utf8");
        const header = (key: string) => new RegExp(`^// ${key}: (.*)$`, "m").exec(text)?.[1]?.trim() ?? "";
        const outcome = header("outcome") === "fail_app_defect" ? "fail_app_defect" : "pass";
        entries.push({ caseId, title: header("title").slice(0, 200), outcome, tested: header("tested"), chars: text.length });
      } catch { /* unreadable file: left out of the list */ }
    }
    return entries.sort((a, b) => a.caseId.localeCompare(b.caseId, undefined, { numeric: true }));
  }

  /** Every system that has scripts, for `excel-partner scripts` with no address. */
  sites(): string[] {
    try { return fs.readdirSync(this.dir, { withFileTypes: true }).filter(item => item.isDirectory()).map(item => item.name).sort(); } catch { return []; }
  }

  read(origin: string, caseId: string): string {
    try { return fs.readFileSync(this.file(origin, caseId), "utf8"); } catch (error) {
      if (error instanceof TestScriptError) throw error;
      throw new TestScriptError(`No script for '${caseId}' on this site.`);
    }
  }

  validate(origin: string, input: { caseId: string; script: string }): void {
    this.file(origin, input.caseId);
    validateScript(origin, input.script);
    const existing = this.list(origin);
    if (!existing.some(item => item.caseId === input.caseId) && existing.length >= MAX_SCRIPTS_PER_SITE) {
      throw new TestScriptError(`This site already has ${MAX_SCRIPTS_PER_SITE} scripts. Delete some (excel-partner scripts delete) before adding more.`);
    }
  }

  save(origin: string, input: { caseId: string; title: string; sheet: string; row: number; outcome: ScriptOutcome; evidence: string; script: string }, today = new Date()): TestScriptEntry {
    this.validate(origin, input);
    const target = this.file(origin, input.caseId);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    const tested = today.toISOString().slice(0, 10);
    const title = oneLine(input.title);
    const header = [
      "// Excel Partner regression script. Written by the assistant; read it before you run it.",
      `// case: ${input.caseId}`,
      `// title: ${title}`,
      `// outcome: ${input.outcome}`,
      `// tested: ${tested}`,
      `// sheet: ${oneLine(input.sheet)} (row ${input.row})`,
      `// system: ${normalizeOrigin(origin)}`,
      `// evidence: ${oneLine(input.evidence).slice(0, 600)}`,
      ...(input.outcome === "fail_app_defect" ? ["// This script asserts the EXPECTED behaviour, so it is supposed to fail until the application defect is fixed."] : []),
      "// Run: npm i -D @playwright/test, then BASE_URL=<system address> STORAGE_STATE=<saved session file> npx playwright test <this file>",
      "",
    ].join("\n");
    fs.writeFileSync(`${target}.tmp`, `${header}${input.script.trim()}\n`, { mode: 0o600 });
    fs.renameSync(`${target}.tmp`, target);
    return { caseId: input.caseId, title: title.slice(0, 200), outcome: input.outcome, tested, chars: input.script.length };
  }

  delete(origin: string, caseId: string): boolean {
    try { fs.rmSync(this.file(origin, caseId)); return true; } catch (error) { if (error instanceof TestScriptError) throw error; return false; }
  }
}
