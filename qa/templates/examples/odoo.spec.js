// TC: TC_000 | Req: -
import { test, expect } from '../_harness/fixtures.js';

test.use({ role: '{{ROLE}}' });

test('[TC_000] The Odoo web client loads for a signed-in user', async ({ page, odoo }) => {
  await odoo.goto('/web');
  await expect(page.locator('.o_web_client')).toBeVisible();
  await odoo.expectNoError();
});
