/** Real Chromium + real bridge/UI, with deterministic providers and isolated settings. */
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { chromium } from 'playwright-core';
import { Bridge } from '../dist/bridge/src/server.js';
import { loadConfig } from '../dist/bridge/src/config.js';
import { MockAdapter } from '../dist/providers/src/mock.js';
import { ProviderRegistry } from '../dist/providers/src/registry.js';

const temporary = await mkdtemp(path.join(os.tmpdir(), 'office-browser-'));
const output = path.resolve(process.env.BROWSER_TEST_OUTPUT ?? 'output/playwright/office-review');
await mkdir(output, { recursive: true });
const config = loadConfig({ host: '127.0.0.1', port: 0, tls: null, publicBaseUrl: 'http://127.0.0.1',
  logLevel: 'error', configFile: path.join(temporary, 'config.json'), skillDirs: [], referenceDirs: [] });
const registry = new ProviderRegistry();
for (const id of ['claude', 'copilot', 'codex', 'opencode']) {
  registry.register(new MockAdapter({ id, reply: 'WEB_SMOKE_OK', delayMs: 150 }));
}
const bridge = new Bridge(config, { registry, usageFile: null, updater: null,
  systemPromptFile: path.join(temporary, 'prompt.txt'), savedReferencesFile: path.join(temporary, 'references.json'),
  workspaceConfigPath: null, skillUploadDir: path.join(temporary, 'skills') });
let browser;
const checks = [], errors = [], failedRequests = [];
const check = async (name, run) => {
  try { await run(); checks.push({ name, ok: true }); console.log(`PASS ${name}`); }
  catch (error) { checks.push({ name, ok: false, error: error.message }); console.error(`FAIL ${name}: ${error.message}`); }
};
try {
  const { url } = await bridge.listen();
  config.publicBaseUrl = url;
  browser = await chromium.launch({ headless: true,
    ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : {}) });
  const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
  page.setDefaultTimeout(10000);
  page.on('pageerror', error => errors.push(error.message));
  page.on('requestfailed', request => failedRequests.push({ url: request.url(), error: request.failure()?.errorText }));
  const settings = async () => { await page.locator('#menu-button').click(); await page.locator('#menu-settings').click(); };
  const connected = () => page.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
  const send = async text => {
    await page.locator('#prompt-input').fill(text);
    await page.locator('#send-button').click();
    await page.getByText('WEB_SMOKE_OK', { exact: true }).waitFor();
  };
  await check('index.html boots and connects', async () => { await page.goto(`${url}/index.html`); await connected(); });
  await check('chat through bridge and deterministic provider', () => send('Browser review message'));
  await check('save instructions and switch theme', async () => {
    await settings();
    await page.locator('#settings-instructions').fill('Reply in Vietnamese.');
    await page.locator('#settings-save-instructions').click();
    await page.waitForFunction(() => document.querySelector('#settings-instructions-status')?.dataset.state === 'saved');
    await page.locator('#theme-select').selectOption('dark');
    assert.equal(await page.locator('html').getAttribute('data-theme'), 'dark');
  });
  await check('switch provider without leaving Settings', async () => {
    await page.locator('#provider-select').selectOption('codex');
    assert.equal(await page.locator('#panel-settings').isVisible(), true);
    await page.locator('[data-close="settings"]').click();
    await send('Second provider message');
  });
  await check('reload preserves conversations in History', async () => {
    await page.reload(); await connected();
    await page.locator('#history-button').click();
    await page.locator('.history-item').filter({ hasText: 'Browser review message' }).waitFor();
    await page.locator('[data-close="history"]').click();
  });
  await check('Settings child preserves draft, scroll and back navigation', async () => {
    await settings();
    await page.locator('#settings-instructions').fill('An unsaved draft');
    await page.locator('#menu-references').scrollIntoViewIfNeeded();
    const scroll = await page.locator('.settings-body').evaluate(el => el.scrollTop);
    await page.locator('#menu-references').click();
    await page.locator('#panel-references [data-close]').click();
    assert.equal(await page.locator('#settings-instructions').inputValue(), 'An unsaved draft');
    assert.ok(Math.abs(await page.locator('.settings-body').evaluate(el => el.scrollTop) - scroll) <= 2);
  });
  await check('Settings and Testing fit desktop and narrow panes in both themes', async () => {
    for (const width of [1200, 380, 320]) {
      await page.setViewportSize({ width, height: 900 });
      for (const theme of ['light', 'dark']) {
        await page.locator('#theme-select').selectOption(theme);
        await page.locator('.settings-body').evaluate(el => { el.scrollTop = 0; });
        for (const panel of ['settings', 'testing']) {
          if (panel === 'testing') await page.locator('#menu-testing').click();
          const bounds = await page.locator(`#panel-${panel}`).evaluate(el => ({ width: el.clientWidth, scroll: el.scrollWidth }));
          assert.ok(bounds.scroll <= bounds.width + 2, `${panel} overflow at ${width}`);
          if (panel === 'testing') {
            const box = await page.locator('#panel-testing .panel-body').boundingBox();
            assert.ok(box.width <= 672, 'Testing should stay compact on wide windows');
            await page.locator('#testing-signin-name').fill('staging-admin');
            await page.locator('#testing-signin-url').fill('https://staging.example.com');
            // Do not submit: that would open an external sign-in browser.
          }
          await page.screenshot({ path: path.join(output, `${panel}-${width}-${theme}.png`), fullPage: true });
          if (panel === 'testing') await page.locator('#panel-testing [data-close]').click();
        }
      }
    }
    await page.locator('[data-close="settings"]').click();
  });
  await check('root route also boots', async () => { await page.goto(url); await connected(); });
  await check('no uncaught browser errors or failed requests', async () => { assert.deepEqual(errors, []); assert.deepEqual(failedRequests, []); });
} finally {
  await browser?.close();
  await bridge.close();
  await rm(temporary, { recursive: true, force: true });
  await writeFile(path.join(output, 'results.json'), JSON.stringify({ checks, errors, failedRequests }, null, 2));
}
if (checks.some(check => !check.ok)) process.exitCode = 1;
console.log(`Browser artifacts: ${output}`);
