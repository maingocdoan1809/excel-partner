import { describe, expect, it } from "vitest";
import path from "node:path";
import fs from "node:fs";
import { validateSource } from "../qa/src/validator.js";
import { allowedOriginsOf } from "../qa/src/config.js";
import { BASE_URL, GOOD_TEST, makeProject } from "./qa-helpers.js";

const { dir, config } = makeProject();
const automationDir = path.join(dir, "automation");
const check = (code: string, extra: Partial<Parameters<typeof validateSource>[1]> = {}) =>
  validateSource(code, { kind: "test", automationDir, file: path.join(automationDir, "tests", "TC_001_x.spec.js"), allowedOrigins: allowedOriginsOf(config), ...extra });
const rules = (code: string) => check(code).violations.map(violation => violation.rule);
const wrap = (body: string, head = "// TC: TC_001 | Req: REQ_1\nimport { test, expect } from '../_harness/fixtures.js';\n") =>
  `${head}test('[TC_001] x', async ({ page, odoo }) => {\n${body}\n  expect(1).toBe(1);\n});\n`;

describe("QA script validator", () => {
  it("accepts a well-formed script and reads its header", () => {
    const result = check(GOOD_TEST);
    expect(result.violations).toEqual([]);
    expect(result.ok).toBe(true);
    expect(result.tcId).toBe("TC_001");
    expect(result.reqId).toBe("REQ_001");
  });

  it("requires the header, a matching file name and titles that start with the id", () => {
    expect(rules(GOOD_TEST.replace("// TC: TC_001 | Req: REQ_001\n", ""))).toContain("header");
    expect(rules(GOOD_TEST.replace("[TC_001]", "[TC_999]"))).toContain("title");
    expect(rules(GOOD_TEST.replace("Sales app opens", "x").replace("[TC_001] x", "no id here"))).toContain("title");
    expect(check(GOOD_TEST, { file: path.join(automationDir, "tests", "TC_002_x.spec.js") }).violations.map(v => v.rule)).toContain("filename");
    expect(check(GOOD_TEST, { expectedTcId: "TC_777" }).violations.map(v => v.rule)).toContain("header");
  });

  it("requires a test and an assertion", () => {
    expect(rules("// TC: TC_001 | Req: -\nimport { test } from '../_harness/fixtures.js';\n")).toContain("no-test");
    expect(rules("// TC: TC_001 | Req: -\nimport { test } from '../_harness/fixtures.js';\ntest('[TC_001] a', async ({ page }) => { await page.goto('/'); });\n")).toContain("no-expect");
  });

  it("rejects TypeScript and broken syntax with a readable message", () => {
    const result = check("// TC: TC_001 | Req: -\nconst x: number = 1;\n");
    expect(result.violations[0]!.rule).toBe("syntax");
    expect(result.violations[0]!.message).toContain("plain JavaScript");
  });

  describe("imports", () => {
    it.each([
      ["node:fs", "import fs from 'node:fs';"],
      ["child_process", "import { exec } from 'child_process';"],
      ["a package", "import axios from 'axios';"],
      ["a path outside the project", "import x from '../../../secret.js';"],
      ["a sibling test", "import x from './TC_002_other.spec.js';"],
    ])("blocks %s", (_label, line) => {
      expect(rules(wrap("", `// TC: TC_001 | Req: -\n${line}\nimport { test, expect } from '../_harness/fixtures.js';\n`))).toContain("import");
    });

    it("blocks dynamic import and import.meta", () => {
      expect(rules(wrap("await import('node:fs');"))).toContain("import");
      expect(rules(wrap("const u = import.meta.url;"))).toContain("import");
    });

    it("allows @playwright/test and the harness, and rejects missing files", () => {
      expect(check(wrap("", "// TC: TC_001 | Req: -\nimport { expect } from '@playwright/test';\nimport { test } from '../_harness/fixtures.js';\n")).violations).toEqual([]);
      expect(rules(wrap("", "// TC: TC_001 | Req: -\nimport { Nope } from '../pages/nope.js';\nimport { test, expect } from '../_harness/fixtures.js';\n"))).toContain("import");
    });
  });

  describe("escape hatches", () => {
    it.each([
      "eval('1')", "new Function('return 1')()", "require('fs')", "process.env.X", "globalThis.x = 1", "await fetch('/x')",
      "new WebSocket('wss://x')", "new XMLHttpRequest()", "const w = new Worker('x')", "Reflect.ownKeys(page)", "new Proxy({}, {})",
    ])("blocks %s", code => { expect(rules(wrap(code))).toContain("identifier"); });

    it.each([
      "await page.route('**/*', r => r.continue())", "await page.context().route('**', r => r.continue())", "await page.unroute('**')",
      "const c = await browser.newContext()", "await chromium.launch()", "const k = await page.context().cookies()",
      "await page.context().addCookies([])", "const s = await page.context().storageState()", "await page.setExtraHTTPHeaders({})",
      "await page.exposeFunction('f', () => 1)", "await page.addInitScript('1')", "const r = page.request", "const f = page.constructor",
      "x.__proto__.y = 1", "await page.context().newCDPSession(page)",
    ])("blocks %s", code => { expect(rules(wrap(code)).some(rule => rule === "member" || rule === "identifier")).toBe(true); });

    it("blocks the request and browser fixtures", () => {
      expect(check(`// TC: TC_001 | Req: -\nimport { test, expect } from '../_harness/fixtures.js';\ntest('[TC_001] a', async ({ request }) => { expect(1).toBe(1); });\n`).violations.map(v => v.rule)).toContain("fixture");
      expect(check(`// TC: TC_001 | Req: -\nimport { test, expect } from '../_harness/fixtures.js';\ntest('[TC_001] a', async ({ browser }) => { expect(1).toBe(1); });\n`).violations.map(v => v.rule)).toContain("fixture");
    });

    it("blocks computed access that can rebuild a banned name", () => {
      expect(rules(wrap("const f = page['con' + 'structor'];"))).toContain("computed-access");
      expect(rules(wrap("const f = page[`con${'structor'}`];"))).toContain("computed-access");
      expect(rules(wrap("const f = page['constructor'];"))).toContain("member");
      expect(rules(wrap("const a = [1,2]; const i = 1; const b = a[0] + a[i] + a[i - 1];"))).toEqual([]);
    });

    it("does not confuse property names or labels with banned globals", () => {
      expect(rules(wrap("const o = { process: 1, fetch: 2 }; const label = o.process;"))).toEqual([]);
    });

    it("blocks .only and debugger", () => {
      expect(rules(wrap("test.only('[TC_001] y', async () => {});"))).toContain("only");
      expect(rules(wrap("debugger;"))).toContain("debugger");
    });
  });

  describe("credentials", () => {
    it("blocks literal passwords typed into password fields", () => {
      expect(rules(wrap("await page.getByLabel('Password').fill('hunter2');"))).toContain("password-literal");
      expect(rules(wrap("await page.locator('input[name=password]').fill(`abc`);"))).toContain("password-literal");
      expect(rules(wrap("await page.getByLabel('Mật khẩu').fill('abc');"))).toContain("password-literal");
    });

    it("allows invalidPassword() and ordinary fields", () => {
      expect(rules(wrap("await page.getByLabel('Password').fill(invalidPassword());\nawait page.getByLabel('Name').fill('Alice');",
        "// TC: TC_001 | Req: -\nimport { test, expect, invalidPassword } from '../_harness/fixtures.js';\n"))).toEqual([]);
    });

    it.each([
      ["a JWT", "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.abcdefghijkl123456"],
      ["a bearer token", "Bearer abcdefghijklmnop1234567890"],
      ["an AWS key", "AKIAABCDEFGHIJKLMNOP"],
      ["a GitHub token", "ghp_abcdefghijklmnopqrstuvwxyz0123456789"],
      ["an opaque key", "k3J9s8dKf02LmZx81Qp7Vb4NcR6yTuHw5AeG0iOd9XzPq"],
    ])("blocks %s in a string", (_label, secret) => { expect(rules(wrap(`const t = '${secret}';`))).toContain("secret"); });

    it("allows ordinary long selectors that contain spaces or brackets", () => {
      expect(rules(wrap("await page.locator('.o_form_view .o_field_widget[name=partner_id] input').first().click();"))).toEqual([]);
    });
  });

  describe("origins and test.use", () => {
    it("blocks absolute URLs outside the allowed origins but allows the configured one and relative paths", () => {
      expect(rules(wrap("await page.goto('https://evil.example.org/x');"))).toContain("origin");
      expect(rules(wrap("await page.goto('//evil.example.org/x');"))).toContain("origin");
      expect(rules(wrap(`await page.goto('${BASE_URL}/web');`))).toEqual([]);
      expect(rules(wrap("await page.goto('/web#action=1');"))).toEqual([]);
    });

    it("only allows harmless test.use options", () => {
      const head = (use: string) => `// TC: TC_001 | Req: -\nimport { test, expect } from '../_harness/fixtures.js';\ntest.use(${use});\n`;
      expect(rules(wrap("", head("{ role: 'admin', viewport: { width: 800, height: 600 } }")))).toEqual([]);
      expect(rules(wrap("", head("{ storageState: 'x.json' }")))).toContain("test-use");
      expect(rules(wrap("", head("{ baseURL: 'https://x.test' }")))).toContain("test-use");
      expect(rules(wrap("", head("{ extraHTTPHeaders: {} }")))).toContain("test-use");
      expect(rules(wrap("", head("someOptions")))).toContain("test-use");
    });
  });

  describe("page objects", () => {
    it("validates imported pages recursively and reports the dependency list", () => {
      const pageFile = path.join(automationDir, "pages", "ok-page.js");
      fs.writeFileSync(pageFile, "export class OkPage { constructor(page) { this.page = page; } }\n");
      const code = wrap("new OkPage(page);", "// TC: TC_001 | Req: -\nimport { test, expect } from '../_harness/fixtures.js';\nimport { OkPage } from '../pages/ok-page.js';\n");
      const good = check(code);
      expect(good.violations).toEqual([]);
      expect(good.deps).toEqual(["automation/pages/ok-page.js"]);

      fs.writeFileSync(pageFile, "import fs from 'node:fs';\nexport class OkPage {}\n");
      const bad = check(code);
      expect(bad.ok).toBe(false);
      expect(bad.violations[0]!.message).toContain("automation/pages/ok-page.js");
    });

    it("survives circular imports between pages", () => {
      fs.writeFileSync(path.join(automationDir, "pages", "a.js"), "import { B } from './b.js';\nexport class A { b = B; }\n");
      fs.writeFileSync(path.join(automationDir, "pages", "b.js"), "import { A } from './a.js';\nexport class B { a = A; }\n");
      const code = wrap("", "// TC: TC_001 | Req: -\nimport { test, expect } from '../_harness/fixtures.js';\nimport { A } from '../pages/a.js';\n");
      expect(check(code).deps).toEqual(["automation/pages/a.js", "automation/pages/b.js"]);
    });
  });

  it("rejects oversized scripts", () => {
    expect(check(`// TC: TC_001 | Req: -\n// ${"x".repeat(61_000)}\n`).violations[0]!.rule).toBe("size");
  });
});
