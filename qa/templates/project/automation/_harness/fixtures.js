// Excel Partner QA harness (managed file).
// Every test builds its browser context here: saved session, origin allowlist, no service workers.
import fs from 'node:fs';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { test as base, expect } from '@playwright/test';
import { isAllowedUrl, allowedOrigins } from './launch.js';
import { odooHelpers } from './odoo.js';
import { collectPageState } from './page-state.js';

export { expect };

/** A password that is certainly wrong, for negative sign-in tests. */
export const invalidPassword = () => `wrong-${randomBytes(6).toString('hex')}`;

export const test = base.extend({
  /** Which saved session to use: test.use({ role: 'admin' }). */
  role: ['default', { option: true }],

  context: async ({ browser, role }, use) => {
    const stateFile = path.join(process.env.QA_AUTH_DIR || 'auth', `${role}.json`);
    if (!fs.existsSync(stateFile)) {
      throw new Error(`No saved session for role "${role}". Run: excel-partner qa login ${role}`);
    }
    const origins = allowedOrigins();
    const context = await browser.newContext({
      baseURL: process.env.QA_BASE_URL,
      storageState: stateFile,
      ignoreHTTPSErrors: process.env.QA_IGNORE_HTTPS_ERRORS === '1',
      serviceWorkers: 'block',
      viewport: { width: 1440, height: 900 },
    });
    // Second layer behind the browser's host rules: nothing outside the allowlist leaves the page.
    await context.route('**/*', route => {
      if (isAllowedUrl(route.request().url(), origins)) return route.continue();
      return route.abort('blockedbyclient');
    });
    await use(context);
    await context.close();
  },

  odoo: async ({ page }, use) => { await use(odooHelpers(page)); },

  /** When a test fails, keep what the page looked like (dialogs, invalid fields, messages). Never changes a result. */
  pageStateOnFailure: [async ({ page }, use, testInfo) => {
    await use();
    if (testInfo.status === testInfo.expectedStatus || testInfo.status === 'skipped') return;
    const state = await collectPageState(page);
    if (state) await testInfo.attach('page-state', { body: JSON.stringify(state), contentType: 'application/json' }).catch(() => {});
  }, { auto: true }],
});
