import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { initProject } from "../qa/src/scaffold.js";
import { loadConfig, type QaConfig } from "../qa/src/config.js";

export const BASE_URL = "https://staging.example.test";

export function makeProject(preset: "odoo" | "generic" = "odoo", baseUrl = BASE_URL): { dir: string; config: QaConfig } {
  const dir = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "qa-project-")), "proj");
  initProject({ dir, preset, baseUrl, install: false });
  return { dir, config: loadConfig(dir) };
}

export const GOOD_TEST = `// TC: TC_001 | Req: REQ_001
import { test, expect } from '../_harness/fixtures.js';

test.use({ role: 'admin' });

test('[TC_001] Sales app opens', async ({ page, odoo }) => {
  await odoo.goto('/web');
  await expect(page.locator('.o_web_client')).toBeVisible();
});
`;
