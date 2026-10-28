import { describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { SessionManager } from "../bridge/src/session-manager.js";
import { ToolGateway } from "../bridge/src/tool-gateway.js";
import { runQaCli, type CliContext } from "../qa/src/cli.js";
import { lineDiff } from "../qa/src/diff.js";
import { QA_SKILLS, installSkills } from "../qa/src/harness.js";
import { MAX_HEAL_ATTEMPTS, approve, listScripts, loadManifest, readApprovedSnapshot, recordRun, resetHeal, resolveScript, saveScript } from "../qa/src/manifest.js";
import { sha256Text } from "../qa/src/util.js";
import { compareAssertions, extractAssertions } from "../qa/src/validator.js";
import { GOOD_TEST, makeProject } from "./qa-helpers.js";

const failing = (dir: string, tcId = "TC_001") => recordRun(dir, `2026102${Math.floor(Math.random() * 9)}-101500`, [{ tcId, status: "failed", durationMs: 10, health: "failing" }]);
const passing = (dir: string, tcId = "TC_001") => recordRun(dir, "20261029-101500", [{ tcId, status: "passed", durationMs: 10, health: "ok" }]);
const save = (dir: string, config: ReturnType<typeof makeProject>["config"], code: string, dryRun = false) => saveScript(dir, config, { tcId: "TC_001", title: "x", code, role: "admin" }, { dryRun });
const reached = (code: string, locator: string) => code.replace(".o_web_client", locator);

describe("what a script checks", () => {
  const script = (body: string) => `// TC: TC_001 | Req: -\nimport { test, expect } from '../_harness/fixtures.js';\ntest('[TC_001] a', async ({ page }) => {\n${body}\n});\n`;

  it("lists each check without the thing it is applied to", () => {
    const found = extractAssertions(script(`
      await expect(page.locator('#a')).toHaveText('Sales Order');
      await expect(page.getByRole('button', { name: 'Ok' })).not.toBeVisible();
      await expect.soft(page).toHaveURL(/web/);
      await expect.poll(async () => 1).toBe(1);
      expect(1 + 1).toBe(2);`));
    expect(found.map(item => item.signature)).toEqual(["toHaveText('Sales Order')", "not.toBeVisible()", "toHaveURL(/web/)", "toBe(1)", "toBe(2)"]);
    expect(found[0]!.subject).toBe("page.locator('#a')");
  });

  it("ignores things that only look like checks", () => {
    expect(extractAssertions(script("await page.locator('#a').click(); const x = expect; await page.getByText('expect(1).toBe(1)').click();"))).toEqual([]);
    expect(extractAssertions("this is not javascript {{{")).toEqual([]);
  });

  it("treats a changed locator as the same check, and a changed expectation as a different one", () => {
    const old = script("await expect(page.locator('.old')).toHaveText('Quotation');\nawait expect(page.locator('.x')).toBeVisible();");
    const moved = script("await expect(page.locator('.new')).toHaveText('Quotation');\nawait expect(page.locator('.x')).toBeVisible();");
    expect(compareAssertions(old, moved)).toEqual({ oldCount: 2, newCount: 2, removed: [], added: [] });
    const reworded = script("await expect(page.locator('.old')).toHaveText('Sales Order');\nawait expect(page.locator('.x')).toBeVisible();");
    expect(compareAssertions(old, reworded)).toMatchObject({ removed: ["toHaveText('Quotation')"], added: ["toHaveText('Sales Order')"] });
    const weaker = script("await expect(page.locator('.old')).toBeVisible();\nawait expect(page.locator('.x')).toBeVisible();");
    expect(compareAssertions(old, weaker).removed).toEqual(["toHaveText('Quotation')"]);
  });

  it("counts repeats and ignores order and extra checks", () => {
    const old = script("await expect(a).toBeVisible();\nawait expect(b).toBeVisible();\nawait expect(c).toHaveCount(3);");
    expect(compareAssertions(old, script("await expect(c).toHaveCount(3);\nawait expect(b).toBeVisible();\nawait expect(a).toBeVisible();\nawait expect(d).toBeEnabled();"))).toMatchObject({ removed: [], added: ["toBeEnabled()"] });
    expect(compareAssertions(old, script("await expect(a).toBeVisible();\nawait expect(c).toHaveCount(3);")).removed).toEqual(["toBeVisible()"]);
  });
});

describe("showing what changed", () => {
  it("marks removed and added lines and keeps a little context", () => {
    const before = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"].join("\n");
    const after = ["a", "b", "c", "d", "E", "f", "g", "h", "i", "j"].join("\n");
    expect(lineDiff(before, after, 1)).toEqual(["  d", "- e", "+ E", "  f"]);
    expect(lineDiff(before, before)).toEqual([]);
  });

  it("separates far-apart changes", () => {
    const lines = Array.from({ length: 20 }, (_, i) => `line ${i}`);
    const changed = lines.map((line, i) => (i === 1 || i === 18 ? `${line} CHANGED` : line));
    const diff = lineDiff(lines.join("\n"), changed.join("\n"), 1);
    expect(diff).toContain("  ...");
    expect(diff.filter(line => line.startsWith("+ ")).length).toBe(2);
  });

  it("copes with files that are too large to compare", () => {
    const big = Array.from({ length: 2500 }, (_, i) => `l${i}`).join("\n");
    expect(lineDiff(big, big + "\nmore")[0]).toContain("too large");
  });
});

describe("limits on repairing a script", () => {
  it("does not count the first save, nor an edit to a script that passes", () => {
    const { dir, config } = makeProject();
    expect(save(dir, config, GOOD_TEST).ok).toBe(true);
    passing(dir);
    expect(save(dir, config, reached(GOOD_TEST, ".o_other")).ok).toBe(true);
    expect(loadManifest(dir).scripts.TC_001!.healAttempts).toBe(0);
  });

  it(`allows ${MAX_HEAL_ATTEMPTS} repairs of a failing script and refuses the next one`, () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    failing(dir);
    expect(save(dir, config, reached(GOOD_TEST, ".o_one")).ok).toBe(true);
    failing(dir);
    expect(save(dir, config, reached(GOOD_TEST, ".o_two")).ok).toBe(true);
    failing(dir);
    const third = save(dir, config, reached(GOOD_TEST, ".o_three"));
    expect(third.ok).toBe(false);
    expect(third.violations[0]).toMatchObject({ rule: "repair-limit" });
    expect(third.violations[0]!.message).toContain("excel-partner qa heal-reset TC_001");
    expect(save(dir, config, reached(GOOD_TEST, ".o_three"), true).ok).toBe(false); // the check also applies before anything is asked
    expect(fs.readFileSync(path.join(dir, loadManifest(dir).scripts.TC_001!.file), "utf8")).toContain(".o_two");
    expect(listScripts(dir, config)[1]!.healAttempts).toBe(MAX_HEAL_ATTEMPTS);
  });

  it("starts over when the script passes, when a person approves, or when a person resets", () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    failing(dir);
    save(dir, config, reached(GOOD_TEST, ".o_one"));
    expect(loadManifest(dir).scripts.TC_001!.healAttempts).toBe(1);
    passing(dir);
    expect(loadManifest(dir).scripts.TC_001!.healAttempts).toBe(0);

    failing(dir);
    save(dir, config, reached(GOOD_TEST, ".o_two"));
    save(dir, config, reached(GOOD_TEST, ".o_three"));
    expect(save(dir, config, reached(GOOD_TEST, ".o_four")).ok).toBe(false);
    resetHeal(dir, "TC_001");
    expect(save(dir, config, reached(GOOD_TEST, ".o_four")).ok).toBe(true);

    approve(dir, config, "TC_001", "person");
    expect(loadManifest(dir).scripts.TC_001!.healAttempts).toBe(0);
  });

  it("treats changing an approved script as a repair", () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    approve(dir, config, "TC_001", "person");
    save(dir, config, reached(GOOD_TEST, ".o_one"));
    save(dir, config, reached(GOOD_TEST, ".o_two"));
    expect(save(dir, config, reached(GOOD_TEST, ".o_three")).violations[0]!.rule).toBe("repair-limit");
    expect(resolveScript(dir, config, "TC_001").status).toBe("modified");
  });

  it("refuses a repair that drops or changes a check, but lets a normal edit through", () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    const weaker = GOOD_TEST.replace(".toBeVisible()", ".toBeAttached()");

    // Not a repair yet (never approved, never failed): a person asked for a different check.
    const edit = save(dir, config, weaker, true);
    expect(edit.ok).toBe(true);

    failing(dir);
    const repair = save(dir, config, weaker);
    expect(repair.ok).toBe(false);
    expect(repair.violations[0]).toMatchObject({ rule: "assertions" });
    expect(repair.violations[0]!.message).toContain("toBeVisible()");
    expect(fs.readFileSync(path.join(dir, loadManifest(dir).scripts.TC_001!.file), "utf8")).toContain("toBeVisible()");
    // A new locator with the same check is fine.
    expect(save(dir, config, reached(GOOD_TEST, ".o_new")).ok).toBe(true);
  });
});

describe("approving again", () => {
  it("keeps the approved text so the next approval can show the difference", () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    expect(readApprovedSnapshot(dir, "TC_001")).toBeNull();
    approve(dir, config, "TC_001", "person");
    expect(readApprovedSnapshot(dir, "TC_001")).toBe(GOOD_TEST);
    expect(fs.existsSync(path.join(dir, "automation", ".approved", "TC_001.spec.js.txt"))).toBe(true);
  });

  function context(cwd: string, answers: string[] = [], interactive = false) {
    const out: string[] = [], err: string[] = [];
    const ctx: CliContext = {
      cwd, configFile: path.join(fs.mkdtempSync(path.join(os.tmpdir(), "qa-repair-cfg-")), "config.json"), out: line => out.push(line), err: line => err.push(line),
      isInteractive: interactive, promptLine: async () => answers.shift() ?? "", promptHidden: async () => "",
    };
    return { ctx, text: () => out.join("\n"), errors: () => err.join("\n") };
  }
  // A person edits the approved file by hand. (Going through the save tool would refuse a repair that drops a check.)
  const approvedThenChanged = (change: (code: string) => string) => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    approve(dir, config, "TC_001", "person");
    fs.writeFileSync(path.join(dir, loadManifest(dir).scripts.TC_001!.file), change(GOOD_TEST));
    return { dir, config };
  };

  it("shows only what changed when a script that was approved before is approved again", async () => {
    const { dir } = approvedThenChanged(code => reached(code, ".o_new"));
    const c = context(dir, ["y"], true);
    expect(await runQaCli(["approve", "TC_001"], c.ctx)).toBe(0);
    expect(c.text()).toContain("Changes since the version you approved");
    expect(c.text()).toContain("- ");
    expect(c.text()).toContain(".o_new");
    expect(c.text()).not.toContain("WARNING");
    expect(c.text()).not.toContain("import { test, expect }"); // the unchanged lines are not repeated
    expect(resolveScript(dir, makeProject().config, "TC_001", loadManifest(dir)).status).toBe("approved");
  });

  it("warns when a check was dropped and does not accept --yes for it", async () => {
    const { dir } = approvedThenChanged(code => code.replace(".toBeVisible()", ".toBeAttached()"));
    const quiet = context(dir);
    expect(await runQaCli(["approve", "TC_001", "--yes"], quiet.ctx)).toBe(1);
    expect(quiet.errors()).toContain("checks were removed or changed");
    expect(quiet.errors()).toContain("toBeVisible()");
    expect(listScripts(dir, makeProject().config).find(item => item.tcId === "TC_001")!.status).toBe("modified"); // still not approved
  });

  it("asks for the word accept in a terminal, and cancels on anything else", async () => {
    const { dir } = approvedThenChanged(code => code.replace(".toBeVisible()", ".toBeAttached()"));
    const no = context(dir, ["y"], true);
    expect(await runQaCli(["approve", "TC_001"], no.ctx)).toBe(0);
    expect(no.text()).toContain("WARNING: these checks were removed or changed");
    expect(no.text()).toContain("not approved");
    const yes = context(dir, ["accept"], true);
    expect(await runQaCli(["approve", "TC_001"], yes.ctx)).toBe(0);
    expect(yes.text()).toContain("TC_001: approved.");
  });

  it("lets a person who has reviewed the change approve it with --accept-assertion-changes", async () => {
    const { dir } = approvedThenChanged(code => code.replace(".toBeVisible()", ".toBeAttached()"));
    const c = context(dir);
    expect(await runQaCli(["approve", "TC_001", "--yes", "--accept-assertion-changes"], c.ctx)).toBe(0);
    expect(readApprovedSnapshot(dir, "TC_001")).toContain("toBeAttached");
  });

  it("lists repairs and resets them from the command line", async () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    approve(dir, config, "TC_001", "person");
    expect(save(dir, config, reached(GOOD_TEST, ".o_new")).ok).toBe(true); // the assistant repairs a locator
    const c = context(dir);
    await runQaCli(["list"], c.ctx);
    expect(c.text()).toContain(`[repairs 1/${MAX_HEAL_ATTEMPTS}]`);
    const reset = context(dir);
    expect(await runQaCli(["heal-reset", "TC_001"], reset.ctx)).toBe(0);
    expect(loadManifest(dir).scripts.TC_001!.healAttempts).toBe(0);
    expect(await runQaCli(["heal-reset"], reset.ctx)).toBe(1);
    expect(await runQaCli(["help"], reset.ctx)).toBe(0);
    expect(reset.text()).toContain("heal-reset <TC_ID...>");
  });
});

describe("the assistant's save tool", () => {
  function tool(dir: string) {
    const sessions = new SessionManager({ ttlMs: 60_000 });
    const session = sessions.create("repair", null);
    Object.defineProperty(session, "isConnected", { get: () => true });
    const asked: unknown[] = [];
    session.requestTool = (async (name: string, args: unknown) => { asked.push({ name, args }); return { ok: true, data: { approved: true } }; }) as never;
    const gateway = new ToolGateway(sessions, { maxReadCells: 1, maxWriteCells: 1, maxSheets: 1, maxSelectionCells: 1 }, 1000, undefined, false, null, null, null, undefined, () => dir);
    return { asked, call: (name: string, args: unknown) => gateway.invoke({ sessionId: session.id, tool: name, args }) };
  }

  it("says plainly when the repair limit is reached, without asking the person", async () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    failing(dir);
    save(dir, config, reached(GOOD_TEST, ".o_one"));
    save(dir, config, reached(GOOD_TEST, ".o_two"));
    const { call, asked } = tool(dir);
    const result: any = await call("qa_save_script", { kind: "test", tcId: "TC_001", title: "x", code: reached(GOOD_TEST, ".o_three") });
    expect(result).toMatchObject({ ok: false, error: { code: "unsupported" } });
    expect(result.error.message).toContain("heal-reset");
    expect(asked).toEqual([]);
    const read: any = await call("qa_read_script", { tcId: "TC_001" });
    expect(read.data).toMatchObject({ repairsSoFar: 2, repairsAllowed: 2 });
  });

  it("tells the assistant to keep the checks when a repair drops one", async () => {
    const { dir, config } = makeProject();
    save(dir, config, GOOD_TEST);
    failing(dir);
    const { call, asked } = tool(dir);
    const result: any = await call("qa_save_script", { kind: "test", tcId: "TC_001", title: "x", code: GOOD_TEST.replace(".toBeVisible()", ".toBeAttached()") });
    expect(result).toMatchObject({ ok: false, error: { code: "invalid_arguments" } });
    expect(result.error.message).toContain("may not remove or change a check");
    expect(result.error.detail[0].rule).toBe("assertions");
    expect(asked).toEqual([]);
  });
});

describe("keeping the skills up to date", () => {
  const read = (target: string, name: string) => fs.readFileSync(path.join(target, name, "SKILL.md"), "utf8");

  it("replaces a copy nobody edited, keeps one that was edited, and replaces anything with force", () => {
    const target = fs.mkdtempSync(path.join(os.tmpdir(), "qa-skills-up-"));
    installSkills(false, target);
    // Pretend an older release had been installed: different text, with the stamp that release wrote.
    const older = "---\nname: qa-heal\ndescription: older\n---\nOld text\n";
    fs.writeFileSync(path.join(target, "qa-heal", "SKILL.md"), older);
    fs.writeFileSync(path.join(target, "qa-heal", ".installed-sha256"), sha256Text(older));
    // And a person edited another one.
    fs.appendFileSync(path.join(target, "qa-report", "SKILL.md"), "\nMy own rule.\n");

    const actions = Object.fromEntries(installSkills(false, target).map(item => [item.name, item.action]));
    expect(actions["qa-heal"]).toBe("updated");
    expect(actions["qa-report"]).toBe("kept");
    expect(actions["qa-testcases"]).toBe("unchanged");
    expect(read(target, "qa-heal")).not.toContain("Old text");
    expect(read(target, "qa-report")).toContain("My own rule.");
    expect(installSkills(true, target).find(item => item.name === "qa-report")!.action).toBe("updated");
    expect(read(target, "qa-report")).not.toContain("My own rule.");
  });

  it("keeps a copy that has no stamp (installed by an earlier release) unless forced", () => {
    const target = fs.mkdtempSync(path.join(os.tmpdir(), "qa-skills-old-"));
    for (const name of QA_SKILLS) { fs.mkdirSync(path.join(target, name), { recursive: true }); fs.writeFileSync(path.join(target, name, "SKILL.md"), "from an earlier release"); }
    expect(installSkills(false, target).every(item => item.action === "kept")).toBe(true);
    expect(installSkills(true, target).every(item => item.action === "updated")).toBe(true);
    expect(installSkills(false, target).every(item => item.action === "unchanged")).toBe(true);
  });
});
