// TC: TC_000 | Req: -
import { test, expect } from '../_harness/fixtures.js';

test.use({ role: '{{ROLE}}' });

test('[TC_000] The home page loads for a signed-in user', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('body')).toBeVisible();
});
