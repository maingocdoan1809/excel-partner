import { describe, expect, it } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { parseConfig, loadConfig, allowedOriginsOf, requireEnvironment } from "../qa/src/config.js";
import { approve, listScripts, loadManifest, resolveScript, revoke, saveManifest, savePage, saveScript } from "../qa/src/manifest.js";
import { checkHarness, installHarness } from "../qa/src/harness.js";
import { QaError, redact, resolveInside, scrubbedEnv, slugify, assertTcId } from "../qa/src/util.js";
import { BASE_URL, GOOD_TEST, makeProject } from "./qa-helpers.js";

describe("QA config", () => {
  it("rejects origins with paths, missing environments and unknown keys", () => {
    const base = { defaultEnvironment: "staging", environments: { staging: { baseUrl: "https://a.test" } } };
    expect(() => parseConfig({ ...base, environments: { staging: { baseUrl: "https://a.test/web" } } })).toThrow(/origin only/);
    expect(() => parseConfig({ ...base, defaultEnvironment: "prod" })).toThrow(/not defined/);
    expect(() => parseConfig({ ...base, surprise: 1 })).toThrow(QaError);
    expect(() => parseConfig({ ...base, allowedOrigins: ["ftp://x.test"] })).toThrow(QaError);
  });

  it("only accepts paths on the site itself, so a cookie cannot be sent elsewhere", () => {
    const base = { defaultEnvironment: "staging", environments: { staging: { baseUrl: "https://a.test" } } };
    for (const bad of ["//evil.test/x", "/\\evil.test", "https://evil.test/x", "web/login", ""]) {
      expect(() => parseConfig({ ...base, profiles: { admin: { login: { path: bad } } } }), bad).toThrow(QaError);
      expect(() => parseConfig({ ...base, profiles: { admin: { check: { path: bad } } } }), bad).toThrow(QaError);
    }
    expect(parseConfig({ ...base, profiles: { admin: { login: { path: "/web/login?db=prod_copy" } } } }).profiles.admin!.login.path).toBe("/web/login?db=prod_copy");
  });

  it("collects the environment origins and the extra ones for the allowlist", () => {
    const config = parseConfig({ defaultEnvironment: "staging", environments: { staging: { baseUrl: "https://a.test/" }, uat: { baseUrl: "https://b.test" } }, allowedOrigins: ["https://login.microsoftonline.com"] });
    expect(allowedOriginsOf(config).sort()).toEqual(["https://a.test", "https://b.test", "https://login.microsoftonline.com"]);
    expect(requireEnvironment(config).baseUrl).toBe("https://a.test");
    expect(() => requireEnvironment(config, "nope")).toThrow(/Unknown environment/);
  });

  it("creates an Odoo project config with sensible sign-in defaults", () => {
    const { config } = makeProject("odoo");
    expect(config.profiles.admin!.login.path).toBe("/web/login");
    expect(config.profiles.admin!.check.type).toBe("odoo-session");
    expect(config.run.retries).toBe(1);
  });
});

describe("QA util", () => {
  it("validates test case ids", () => {
    for (const ok of ["TC_001", "TC-SALES-12", "LOGIN_01", "A1_2"]) expect(assertTcId(ok)).toBe(ok);
    for (const bad of ["001", "TC", "TC 001", "../x", "TC_001/../x", "TC__1", ""]) expect(() => assertTcId(bad)).toThrow(QaError);
  });

  it("slugifies titles including Vietnamese", () => {
    expect(slugify("Đăng nhập thành công!")).toBe("dang-nhap-thanh-cong");
    expect(slugify("???")).toBe("test");
  });

  it("keeps paths inside the project", () => {
    const { dir } = makeProject();
    expect(resolveInside(dir, "automation/tests/x.js")).toBe(path.join(dir, "automation", "tests", "x.js"));
    for (const bad of ["../x", "automation/../../x", "/etc/passwd", "C:\\Windows\\x", "a\0b", ""]) expect(() => resolveInside(dir, bad)).toThrow(QaError);
  });

  it("refuses a link that points out of the project", () => {
    const { dir } = makeProject();
    const outside = fs.mkdtempSync(path.join(path.dirname(dir), "outside-"));
    try { fs.symlinkSync(outside, path.join(dir, "automation", "link"), "junction"); } catch { return; }
    expect(() => resolveInside(dir, "automation/link/x.js")).toThrow(/outside/);
  });

  it("scrubs the environment down to what a browser run needs", () => {
    const env = scrubbedEnv({ QA_BASE_URL: "https://a.test" }, {
      PATH: "p", SystemRoot: "C:\\Windows", TEMP: "t", ANTHROPIC_API_KEY: "sk-1", GITHUB_TOKEN: "gh", OPENAI_API_KEY: "o", USERNAME: "u", NODE_OPTIONS: "--require evil.js",
    });
    expect(env).toEqual({ PATH: "p", SystemRoot: "C:\\Windows", TEMP: "t", QA_BASE_URL: "https://a.test" });
    expect(scrubbedEnv({}, { NODE_OPTIONS: "--max-old-space-size=4096" }).NODE_OPTIONS).toBe("--max-old-space-size=4096");
  });

  it("redacts credentials from text", () => {
    const text = [
      "Authorization: Bearer abcdefghijklmnop12345", "Cookie: session_id=abc123; other=1", "password=hunter2&x=1",
      "token eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.abcdefghijkl123456", 'api_key: "sk-live-1234"',
    ].join("\n");
    const out = redact(text);
    for (const secret of ["abcdefghijklmnop12345", "abc123", "hunter2", "eyJhbGciOiJIUzI1NiJ9", "sk-live-1234"]) expect(out).not.toContain(secret);
    expect(out).toContain("x=1");
  });

  it("keeps the harness redactor in sync with the bridge-side one", async () => {
    const harness = await import(`${path.resolve("qa/templates/project/automation/_harness/redact.js")}`.replace(/^([A-Za-z]):/, "file:///$1:").replace(/\\/g, "/"));
    const samples = ["Authorization: Bearer abcdefghijklmnop12345", "Cookie: session_id=abc123", "password=hunter2&x=1", "Bearer abcdefghijklmnopqrstuv", "nothing secret here"];
    for (const sample of samples) expect(harness.redact(sample)).toBe(redact(sample));
  });
});

describe("script lifecycle", () => {
  it("saves a draft, approves it, and voids the approval when the file changes", () => {
    const { dir, config } = makeProject();
    const saved = saveScript(dir, config, { tcId: "TC_001", title: "Sales app opens", code: GOOD_TEST, role: "admin", tags: ["@smoke"] });
    expect(saved).toMatchObject({ ok: true, status: "draft", file: "automation/tests/TC_001_sales-app-opens.spec.js" });
    expect(resolveScript(dir, config, "TC_001").status).toBe("draft");

    approve(dir, config, "TC_001", "tester");
    expect(resolveScript(dir, config, "TC_001").status).toBe("approved");
    expect(loadManifest(dir).scripts.TC_001!.approved!.by).toBe("tester");
    expect(loadManifest(dir).scripts.TC_001!.tags).toEqual(["smoke"]);

    // Editing on disk is detected even though the manifest was not touched.
    const file = path.join(dir, saved.file!);
    fs.writeFileSync(file, fs.readFileSync(file, "utf8") + "\n// sneaky edit\n");
    expect(resolveScript(dir, config, "TC_001").status).toBe("modified");

    // Saving through the tool keeps the old approval attached to the old hash, so it reads as modified.
    const resaved = saveScript(dir, config, { tcId: "TC_001", title: "Sales app opens", code: GOOD_TEST.replace("web'", "web#x'"), role: "admin" });
    expect(resaved.status).toBe("modified");
    expect(listScripts(dir, config).find(item => item.tcId === "TC_001")!.status).toBe("modified");

    approve(dir, config, "TC_001", "tester");
    expect(resolveScript(dir, config, "TC_001").status).toBe("approved");
    revoke(dir, "TC_001");
    expect(resolveScript(dir, config, "TC_001").status).toBe("draft");
  });

  it("returns violations instead of writing an invalid script", () => {
    const { dir, config } = makeProject();
    const result = saveScript(dir, config, { tcId: "TC_050", title: "Bad", code: GOOD_TEST.replace("TC_001", "TC_050").replace("import {", "import fs from 'node:fs';\nimport {"), role: "admin" });
    expect(result.ok).toBe(false);
    expect(result.violations[0]!.rule).toBe("import");
    expect(fs.existsSync(path.join(dir, "automation", "tests", "TC_050_bad.spec.js"))).toBe(false);
    expect(loadManifest(dir).scripts.TC_050).toBeUndefined();
  });

  it("refuses to approve an invalid script even if the manifest says it is fine", () => {
    const { dir, config } = makeProject();
    saveScript(dir, config, { tcId: "TC_001", title: "x", code: GOOD_TEST, role: "admin" });
    const file = path.join(dir, loadManifest(dir).scripts.TC_001!.file);
    fs.writeFileSync(file, fs.readFileSync(file, "utf8").replace("import {", "import fs from 'node:fs';\nimport {"));
    expect(() => approve(dir, config, "TC_001", "x")).toThrow(/fails validation/);
    expect(resolveScript(dir, config, "TC_001").status).toBe("invalid");
  });

  it("voids approvals of tests that use a page object when the page changes", () => {
    const { dir, config } = makeProject();
    savePage(dir, config, "sale-page", "export class SalePage { constructor(page) { this.page = page; } }\n");
    const code = GOOD_TEST.replace("import { test, expect }", "import { SalePage } from '../pages/sale-page.js';\nimport { test, expect }").replace("await odoo.goto('/web');", "new SalePage(page);\n  await odoo.goto('/web');");
    expect(saveScript(dir, config, { tcId: "TC_001", title: "uses page", code, role: "admin" }).ok).toBe(true);
    approve(dir, config, "TC_001", "x");
    expect(resolveScript(dir, config, "TC_001").status).toBe("approved");
    expect(savePage(dir, config, "sale-page", "export class SalePage { constructor(page) { this.page = page; this.changed = true; } }\n").ok).toBe(true);
    expect(resolveScript(dir, config, "TC_001").status).toBe("modified");
  });

  it("rejects unknown roles and page names that could escape the folder", () => {
    const { dir, config } = makeProject();
    expect(() => saveScript(dir, config, { tcId: "TC_001", title: "x", code: GOOD_TEST, role: "ghost" })).toThrow(/Unknown role/);
    expect(() => savePage(dir, config, "../evil", "x")).toThrow(/Page name/);
  });

  it("treats a forged manifest approval as modified", () => {
    const { dir, config } = makeProject();
    saveScript(dir, config, { tcId: "TC_001", title: "x", code: GOOD_TEST, role: "admin" });
    const manifest = loadManifest(dir);
    manifest.scripts.TC_001!.approved = { by: "forger", at: new Date().toISOString(), hash: "0".repeat(64) };
    saveManifest(dir, manifest);
    expect(resolveScript(dir, config, "TC_001").status).toBe("modified");
  });

  it("keeps the example script TC_000 as a draft after init", () => {
    const { dir, config } = makeProject();
    expect(listScripts(dir, config).map(item => [item.tcId, item.status])).toEqual([["TC_000", "draft"]]);
  });
});

describe("harness integrity", () => {
  it("passes after init, flags edits, and is restored by upgrade", () => {
    const { dir } = makeProject();
    expect(checkHarness(dir)).toMatchObject({ ok: true, outdated: false });
    const fixtures = path.join(dir, "automation", "_harness", "fixtures.js");
    fs.appendFileSync(fixtures, "\n// backdoor\n");
    const broken = checkHarness(dir);
    expect(broken.ok).toBe(false);
    expect(broken.problems[0]).toContain("fixtures.js was modified");
    fs.rmSync(path.join(dir, "automation", "_harness", "odoo.js"));
    expect(checkHarness(dir).problems.join(" ")).toContain("odoo.js is missing");
    installHarness(dir);
    expect(checkHarness(dir).ok).toBe(true);
  });

  it("keeps secrets, sessions and results out of git", () => {
    const { dir } = makeProject();
    const ignore = fs.readFileSync(path.join(dir, ".gitignore"), "utf8");
    expect(ignore).toMatch(/^auth\/$/m);
    expect(ignore).toMatch(/^results\/$/m);
    expect(ignore).toMatch(/^\.env$/m);
  });

  it("refuses to initialise over an existing project", async () => {
    const { dir } = makeProject();
    const { initProject } = await import("../qa/src/scaffold.js");
    expect(() => initProject({ dir, preset: "odoo", baseUrl: BASE_URL, install: false })).toThrow(/already a QA project/);
    expect(loadConfig(dir).defaultEnvironment).toBe("staging");
  });
});
