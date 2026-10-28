import { describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { runQaCli, type CliContext } from "../qa/src/cli.js";
import { readActiveProject } from "../qa/src/active.js";
import { loadManifest, saveScript } from "../qa/src/manifest.js";
import { loadConfig } from "../qa/src/config.js";
import { GOOD_TEST } from "./qa-helpers.js";

function context(cwd: string, answers: string[] = [], interactive = false) {
  const out: string[] = [], err: string[] = [];
  const configFile = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "qa-cli-cfg-")), "config.json");
  const ctx: CliContext = {
    cwd, configFile, out: line => out.push(line), err: line => err.push(line), isInteractive: interactive,
    promptLine: async () => answers.shift() ?? "", promptHidden: async () => "pw",
  };
  return { ctx, out, err, text: () => out.join("\n"), errText: () => err.join("\n") };
}

async function newProject() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "qa-cli-"));
  const c = context(root);
  expect(await runQaCli(["init", "proj", "--url", "https://staging.example.test/web", "--no-install"], c.ctx)).toBe(0);
  return { root, dir: path.join(root, "proj"), c };
}

describe("excel-partner qa (CLI)", () => {
  it("prints help and rejects unknown commands", async () => {
    const c = context(os.tmpdir());
    expect(await runQaCli(["help"], c.ctx)).toBe(0);
    expect(c.text()).toContain("approve <TC_ID...>");
    expect(await runQaCli(["banana"], c.ctx)).toBe(2);
  });

  it("init creates the project, strips the URL to its origin and records it as the active project", async () => {
    const { dir, c } = await newProject();
    expect(loadConfig(dir).environments.staging!.baseUrl).toBe("https://staging.example.test");
    expect(readActiveProject(c.ctx.configFile)).toBe(dir);
    expect(c.text()).toContain("can reach nothing else");
    for (const file of ["qa.config.json", ".gitignore", "package.json", "automation/_harness/fixtures.js", "automation/playwright.config.js", "automation/tests/TC_000_application-loads.spec.js"]) {
      expect(fs.existsSync(path.join(dir, file)), file).toBe(true);
    }
  });

  it("init needs a URL and refuses a non-empty folder", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "qa-cli-"));
    const c = context(root);
    expect(await runQaCli(["init", "p", "--no-install"], c.ctx)).toBe(1);
    expect(c.errText()).toContain("--url");
    fs.mkdirSync(path.join(root, "full")); fs.writeFileSync(path.join(root, "full", "x.txt"), "x");
    expect(await runQaCli(["init", "full", "--url", "https://a.test", "--no-install"], c.ctx)).toBe(1);
    expect(c.errText()).toContain("not empty");
  });

  it("lists scripts and validates them, reporting violations with line numbers", async () => {
    const { dir } = await newProject();
    const c = context(dir);
    expect(await runQaCli(["list"], c.ctx)).toBe(0);
    expect(c.text()).toContain("TC_000");
    expect(c.text()).toContain("draft");
    const config = loadConfig(dir);
    saveScript(dir, config, { tcId: "TC_001", title: "x", code: GOOD_TEST, role: "admin" });
    const file = path.join(dir, loadManifest(dir).scripts.TC_001!.file);
    fs.writeFileSync(file, fs.readFileSync(file, "utf8").replace("await odoo.goto('/web');", "await odoo.goto('/web');\n  eval('1');"));
    const v = context(dir);
    expect(await runQaCli(["validate"], v.ctx)).toBe(1);
    expect(v.text()).toContain("TC_000: ok");
    expect(v.text()).toMatch(/TC_001: FAIL[\s\S]*line \d+: 'eval' is not allowed/);
  });

  it("approval needs a person: interactive confirmation or an explicit --yes", async () => {
    const { dir } = await newProject();
    const none = context(dir);
    expect(await runQaCli(["approve", "TC_000"], none.ctx)).toBe(1);
    expect(none.errText()).toContain("needs a person");
    expect(loadManifest(dir).scripts.TC_000!.approved).toBeUndefined();

    const declined = context(dir, ["n"], true);
    expect(await runQaCli(["approve", "TC_000"], declined.ctx)).toBe(0);
    expect(declined.text()).toContain("not approved");
    expect(declined.text()).toContain("test('[TC_000]");
    expect(loadManifest(dir).scripts.TC_000!.approved).toBeUndefined();

    const accepted = context(dir, ["y"], true);
    expect(await runQaCli(["approve", "TC_000"], accepted.ctx)).toBe(0);
    expect(loadManifest(dir).scripts.TC_000!.approved).toBeDefined();

    const again = context(dir);
    expect(await runQaCli(["revoke", "TC_000"], again.ctx)).toBe(0);
    expect(loadManifest(dir).scripts.TC_000!.approved).toBeUndefined();
    expect(await runQaCli(["approve", "TC_000", "--yes"], again.ctx)).toBe(0);
    expect(loadManifest(dir).scripts.TC_000!.approved).toBeDefined();
  });

  it("will not approve a script that fails validation", async () => {
    const { dir } = await newProject();
    const file = path.join(dir, loadManifest(dir).scripts.TC_000!.file);
    fs.writeFileSync(file, fs.readFileSync(file, "utf8") + "\nprocess.exit(1);\n");
    const c = context(dir);
    expect(await runQaCli(["approve", "TC_000", "--yes"], c.ctx)).toBe(1);
    expect(c.errText()).toContain("fails validation");
  });

  it("finds the project from a sub-folder, or through --project, and complains clearly when there is none", async () => {
    const { dir } = await newProject();
    const sub = path.join(dir, "automation", "tests");
    expect(await runQaCli(["list"], context(sub).ctx)).toBe(0);
    expect(await runQaCli(["list", "--project", dir], context(os.tmpdir()).ctx)).toBe(0);
    const nowhere = context(fs.mkdtempSync(path.join(os.tmpdir(), "qa-none-")));
    expect(await runQaCli(["list"], nowhere.ctx)).toBe(1);
    expect(nowhere.errText()).toContain("qa init");
  });

  it("status reports a missing session as invalid", async () => {
    const { dir } = await newProject();
    const c = context(dir);
    expect(await runQaCli(["status"], c.ctx)).toBe(1);
    expect(c.text()).toMatch(/admin\s+INVALID \(no saved session\)/);
  });

  it("doctor flags the missing Playwright install and an unsafe-looking production URL", async () => {
    const { dir } = await newProject();
    const c = context(dir);
    expect(await runQaCli(["doctor"], c.ctx)).toBe(1);
    expect(c.text()).toContain("[fail] Playwright is not installed");
    expect(c.text()).toContain("[ ok ] Harness files intact");
    expect(c.text()).toContain("[ ok ] .gitignore keeps auth/");
  });

  it("run explains how to get a script approved instead of silently doing nothing", async () => {
    const { dir } = await newProject();
    const c = context(dir);
    // No Playwright in this project: the error names the fix.
    expect(await runQaCli(["run"], c.ctx)).toBe(1);
    expect(c.errText()).toContain("Playwright is not installed");
  });
});
