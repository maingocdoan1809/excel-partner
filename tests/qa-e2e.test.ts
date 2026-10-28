import { afterAll, beforeAll, describe, expect, it } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { approve, loadManifest, saveScript } from "../qa/src/manifest.js";
import { installHarness } from "../qa/src/harness.js";
import { runLogin } from "../qa/src/runtime.js";
import { readReport, runTests } from "../qa/src/runner.js";
import { authFile, checkSession } from "../qa/src/session.js";
import { installDependencies } from "../qa/src/scaffold.js";
import { startFakeOdoo, type FakeOdoo } from "./qa-fake-odoo.js";
import { makeProject } from "./qa-helpers.js";
import type { QaConfig } from "../qa/src/config.js";
import { allowedOriginsOf } from "../qa/src/config.js";
import { ExternalTools } from "../bridge/src/external-tools.js";

const script = (id: string, title: string, body: string) => `// TC: ${id} | Req: REQ_${id.slice(-3)}
import { test, expect } from '../_harness/fixtures.js';

test.use({ role: 'admin' });

test('[${id}] ${title}', async ({ page, odoo }) => {
${body}
});
`;
const LOADS = "  await odoo.goto('/web');\n  await expect(page.locator('#hello')).toHaveText('Hello admin');";
const FAILS = "  await odoo.goto('/web');\n  await expect(page.locator('#hello')).toHaveText('Nope', { timeout: 1500 });";
const EGRESS = "  await odoo.goto('/web');\n  await expect(page.locator('#egress')).toHaveText('blocked');\n  await page.waitForTimeout(500);";

let fake: FakeOdoo;
let dir = "";
let config: QaConfig;
let ready = false;

beforeAll(async () => {
  fake = await startFakeOdoo();
  ({ dir, config } = makeProject("odoo", fake.origin));
  try { installDependencies(dir, undefined, true); ready = true; } catch { ready = false; }
  process.env.QA_ADMIN_USER = "admin";
  process.env.QA_ADMIN_PASSWORD = "s3cret";
}, 240_000);

afterAll(async () => {
  delete process.env.QA_ADMIN_USER; delete process.env.QA_ADMIN_PASSWORD;
  await fake?.close();
});

const need = (ctx: { skip: () => void }) => { if (!ready) ctx.skip(); };
const T = 180_000;

describe("QA harness end to end (real Playwright)", () => {
  it("signs in headlessly with credentials supplied through the environment and saves a session the server accepts", async ctx => {
    need(ctx);
    expect((await checkSession(dir, config, "staging", "admin")).reason).toBe("no saved session");
    await runLogin({ projectDir: dir, config, profile: "admin", mode: "headless", credentials: { user: "admin", password: "s3cret" } });
    expect(fs.existsSync(authFile(dir, "staging", "admin"))).toBe(true);
    expect(await checkSession(dir, config, "staging", "admin")).toMatchObject({ valid: true });
    expect(fake.logins).toBe(1);
  }, T);

  it("does not wait for a network that never goes idle (Odoo keeps a long-poll open)", async ctx => {
    need(ctx);
    const started = Date.now();
    await runLogin({ projectDir: dir, config, profile: "admin", mode: "headless", credentials: { user: "admin", password: "s3cret" } });
    expect(Date.now() - started).toBeLessThan(15_000);
  }, T);

  it("reports a wrong password without saving anything", async ctx => {
    need(ctx);
    const before = fs.readFileSync(authFile(dir, "staging", "admin"), "utf8");
    await expect(runLogin({ projectDir: dir, config, profile: "admin", mode: "headless", credentials: { user: "admin", password: "nope" }, timeoutMs: 40_000 })).rejects.toThrow(/Sign-in did not complete/);
    expect(fs.readFileSync(authFile(dir, "staging", "admin"), "utf8")).toBe(before);
  }, T);

  it("runs approved scripts, records results and evidence, and keeps the browser inside the allowlist", async ctx => {
    need(ctx);
    for (const [id, title, body] of [["TC_010", "Web client loads", LOADS], ["TC_011", "Wrong text fails", FAILS], ["TC_012", "Page cannot call out", EGRESS]] as const) {
      expect(saveScript(dir, config, { tcId: id, title, code: script(id, title, body), role: "admin" }).ok).toBe(true);
      approve(dir, config, id, "tester");
    }
    const lines: string[] = [];
    const report = await runTests({ projectDir: dir, onLine: line => lines.push(line) });

    const byId = Object.fromEntries(report.tests.map(test => [test.tcId, test]));
    expect(Object.keys(byId).sort()).toEqual(["TC_010", "TC_011", "TC_012"]);
    expect(byId.TC_010).toMatchObject({ status: "passed", health: "ok", role: "admin", scriptStatus: "approved" });
    expect(byId.TC_012).toMatchObject({ status: "passed" });
    expect(byId.TC_011).toMatchObject({ status: "failed", health: "failing", retries: 1 });
    expect(byId.TC_011!.error).toContain("Nope");
    // The harness keeps what the page looked like at the moment of failure, and where the script stopped.
    expect(byId.TC_011!.pageState?.dialogs[0]).toMatchObject({ title: "Confirmation", buttons: ["Ok", "Cancel"] });
    expect(byId.TC_011!.failedAt?.line).toBeGreaterThan(0);
    expect(byId.TC_011!.failedStep).toBeTruthy();
    expect(byId.TC_010!.pageState).toBeUndefined();
    expect(report.summary).toMatchObject({ total: 3, passed: 2, failed: 1 });

    // A failing test leaves a screenshot inside the run folder; passing ones leave none.
    expect(byId.TC_011!.evidence.some(file => file.endsWith(".png") && fs.existsSync(path.join(dir, file)))).toBe(true);
    expect(byId.TC_010!.evidence).toEqual([]);

    // Nothing reached the non-allowed origin, neither by script nor by image tag.
    expect(fake.leaks).toEqual([]);

    // The manifest remembers the last run; the report can be read back.
    expect(loadManifest(dir).scripts.TC_011!.lastRun).toMatchObject({ runId: report.runId, health: "failing" });
    expect(readReport(dir).runId).toBe(report.runId);
    expect(readReport(dir, report.runId).tests).toHaveLength(3);
  }, T);

  it("never writes the password or session cookie into results", async ctx => {
    need(ctx);
    const state = JSON.parse(fs.readFileSync(authFile(dir, "staging", "admin"), "utf8")) as { cookies: Array<{ value: string }> };
    const cookie = state.cookies[0]!.value;
    const walk = (folder: string): string[] => fs.readdirSync(folder, { withFileTypes: true }).flatMap(entry => entry.isDirectory() ? walk(path.join(folder, entry.name)) : [path.join(folder, entry.name)]);
    for (const file of walk(path.join(dir, "results")).filter(name => /\.(json|txt|md|log)$/.test(name))) {
      const text = fs.readFileSync(file, "utf8");
      expect(text, file).not.toContain("s3cret");
      expect(text, file).not.toContain(cookie);
    }
  }, T);

  it("only runs drafts when they are named, and says why an unapproved script was skipped", async ctx => {
    need(ctx);
    saveScript(dir, config, { tcId: "TC_013", title: "Draft only", code: script("TC_013", "Draft only", LOADS), role: "admin" });
    const sweep = await runTests({ projectDir: dir, tcIds: [], tag: "none" });
    expect(sweep.tests).toEqual([]);
    const named = await runTests({ projectDir: dir, tcIds: ["TC_013"] });
    expect(named.tests).toEqual([]);
    expect(named.notRun[0]).toMatchObject({ tcId: "TC_013" });
    expect(named.notRun[0]!.reason).toContain("not approved");
    await expect(runTests({ projectDir: dir, draft: true })).rejects.toThrow(/name them/);
    const draft = await runTests({ projectDir: dir, tcIds: ["TC_013"], draft: true, headed: false });
    expect(draft.tests[0]).toMatchObject({ tcId: "TC_013", status: "passed", scriptStatus: "draft" });
    expect(draft.draftMode).toBe(true);
  }, T);

  it("refuses to run a script that changed after it was approved", async ctx => {
    need(ctx);
    const file = path.join(dir, loadManifest(dir).scripts.TC_010!.file);
    const original = fs.readFileSync(file, "utf8");
    fs.writeFileSync(file, original + "\n// edited after approval\n");
    try {
      const report = await runTests({ projectDir: dir, tcIds: ["TC_010"] });
      expect(report.tests).toEqual([]);
      expect(report.notRun[0]!.reason).toContain("changed since it was approved");
    } finally { fs.writeFileSync(file, original); }
  }, T);

  it("refuses to run when the harness was tampered with, and runs again after upgrade", async ctx => {
    need(ctx);
    const fixtures = path.join(dir, "automation", "_harness", "fixtures.js");
    fs.appendFileSync(fixtures, "\n// let everything through\n");
    await expect(runTests({ projectDir: dir, tcIds: ["TC_010"] })).rejects.toThrow(/harness was changed/);
    installHarness(dir);
    expect((await runTests({ projectDir: dir, tcIds: ["TC_010"] })).summary.passed).toBe(1);
  }, T);

  it("signs in again by itself when the server forgets the session and a stored password is available", async ctx => {
    need(ctx);
    const before = fake.logins;
    fake.expireSessions();
    expect((await checkSession(dir, config, "staging", "admin")).valid).toBe(false);
    const lines: string[] = [];
    const report = await runTests({ projectDir: dir, tcIds: ["TC_010"], onLine: line => lines.push(line) });
    expect(report.summary.passed).toBe(1);
    expect(fake.logins).toBe(before + 1);
    expect(lines.join("\n")).toContain("signing in again");
  }, T);

  it("stops with the exact command to run when the session expired and no password is stored", async ctx => {
    need(ctx);
    fake.expireSessions();
    const user = process.env.QA_ADMIN_USER, password = process.env.QA_ADMIN_PASSWORD;
    delete process.env.QA_ADMIN_USER; delete process.env.QA_ADMIN_PASSWORD;
    try {
      await expect(runTests({ projectDir: dir, tcIds: ["TC_010"] })).rejects.toThrow(/excel-partner qa login admin/);
    } finally { process.env.QA_ADMIN_USER = user; process.env.QA_ADMIN_PASSWORD = password; }
  }, T);

  it("gives the assistant a signed-in browser that cannot leave the allowed site", async ctx => {
    need(ctx);
    // The session may have been refreshed above; make sure a valid one exists.
    if (!(await checkSession(dir, config, "staging", "admin")).valid) await runLogin({ projectDir: dir, config, profile: "admin", mode: "headless", credentials: { user: "admin", password: "s3cret" } });
    const tools = new ExternalTools();
    try {
      await tools.useProfile("qa-session", { storageState: authFile(dir, "staging", "admin"), allowedOrigins: allowedOriginsOf(config) });
      const text = (result: any) => JSON.stringify(result);
      await tools.browser("qa-session", "browser_navigate", { url: `${fake.origin}/web` });
      // Signed in without ever typing a password: the page shows the user's name.
      expect(text(await tools.browser("qa-session", "browser_snapshot", {}))).toContain("Hello admin");
      // Give the page's own background calls (to the other origin) time to fail, then look at the server.
      await new Promise(resolve => setTimeout(resolve, 1500));
      expect(fake.leaks).toEqual([]);
      const away = await tools.browser("qa-session", "browser_navigate", { url: `${fake.otherOrigin}/direct` }).then(text, error => String(error));
      expect(fake.leaks).toEqual([]);
      expect(away).toMatch(/blocked|not allowed|ERR_|error/i);
      // The saved session file was copied, not modified.
      expect(JSON.parse(fs.readFileSync(authFile(dir, "staging", "admin"), "utf8")).cookies.length).toBeGreaterThan(0);
    } finally { await tools.close(); }
  }, 300_000);
});
