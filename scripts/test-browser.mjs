/** Real Chromium + real bridge/UI, with deterministic providers and isolated settings. */
import { mkdtemp, mkdir, writeFile, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';
import { Bridge } from '../dist/bridge/src/server.js';
import { loadConfig } from '../dist/bridge/src/config.js';
import { MockAdapter } from '../dist/providers/src/mock.js';
import { ProviderRegistry } from '../dist/providers/src/registry.js';

// Node's in-process mock MCP client needs the same local CA as the browser and real MCP children.
if (process.env.BROWSER_TEST_CERT && process.env.BROWSER_TEST_TLS_BOOTSTRAPPED !== '1') {
  const result = spawnSync(process.execPath, [fileURLToPath(import.meta.url)], {
    env: { ...process.env, NODE_EXTRA_CA_CERTS: path.resolve(process.env.BROWSER_TEST_CERT), BROWSER_TEST_TLS_BOOTSTRAPPED: '1' }, stdio: 'inherit',
  });
  process.exit(result.status ?? 1);
}
const temporary = await mkdtemp(path.join(os.tmpdir(), 'office-browser-'));
process.env.OFFICE_PARTNER_DATA_DIR = temporary;
process.env.BRIDGE_CONFIG = path.join(temporary, 'config.json');
const output = path.resolve(process.env.BROWSER_TEST_OUTPUT ?? 'output/playwright/office-review');
await mkdir(output, { recursive: true });
const tls = process.env.BROWSER_TEST_CERT && process.env.BROWSER_TEST_KEY
  ? { certPath: process.env.BROWSER_TEST_CERT, keyPath: process.env.BROWSER_TEST_KEY } : null;
const config = loadConfig({ host: '127.0.0.1', port: Number(process.env.BROWSER_TEST_PORT ?? 0), tls, publicBaseUrl: 'http://127.0.0.1',
  logLevel: 'error', configFile: path.join(temporary, 'config.json'), skillDirs: [], referenceDirs: [] });
const registry = new ProviderRegistry();
const mockCalls = new Map();
for (const id of ['claude', 'copilot', 'codex', 'opencode']) {
  const toolCalls = []; mockCalls.set(id, toolCalls);
  registry.register(new MockAdapter({ id, toolCalls, reply: 'WEB_SMOKE_OK', delayMs: 150 }));
}
const bridge = new Bridge(config, { registry, usageFile: null, updater: null,
  systemPromptFile: path.join(temporary, 'prompt.txt'), savedReferencesFile: path.join(temporary, 'references.json'),
  workspaceConfigPath: null, skillUploadDir: path.join(temporary, 'skills') });
let browser, page;
const checks = [], errors = [], failedRequests = [];
const check = async (name, run) => {
  try { await run(); checks.push({ name, ok: true }); console.log(`PASS ${name}`); }
  catch (error) {
    const file = name.replace(/[^a-z0-9]+/gi, '-').toLowerCase();
    await page?.screenshot({ path: path.join(output, `failure-${file}.png`), fullPage: true }).catch(() => {});
    const state = await page?.evaluate(() => ({ provider: document.querySelector('#provider-select')?.value,
      permission: document.querySelector('#permission-mode')?.value, errors: [...document.querySelectorAll('.bubble.error')].map(el => el.textContent) })).catch(() => undefined);
    checks.push({ name, ok: false, error: error.message, state }); console.error(`FAIL ${name}: ${error.message} ${JSON.stringify(state)}`);
  }
};
try {
  const listening = await bridge.listen();
  const url = process.env.BROWSER_TEST_PORT ? `${tls ? 'https' : 'http'}://localhost:${process.env.BROWSER_TEST_PORT}` : listening.url;
  config.publicBaseUrl = url;
  browser = await chromium.launch({ headless: process.env.BROWSER_TEST_HEADED !== '1',
    ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : process.platform === 'win32' ? { channel: 'msedge' } : {}) });
  page = await browser.newPage({ viewport: { width: 1200, height: 900 }, ignoreHTTPSErrors: Boolean(tls) });
  page.setDefaultTimeout(10000);
  page.on('pageerror', error => errors.push(error.message));
  page.on('requestfailed', request => failedRequests.push({ url: request.url(), error: request.failure()?.errorText }));
  const settings = async () => { await page.locator('#menu-button').click(); await page.locator('#panel-settings').waitFor({ state: 'visible' }); };
  const connected = () => page.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
  const send = async text => {
    await page.locator('#prompt-input').fill(text);
    await page.locator('#send-button').click();
    await page.getByText('WEB_SMOKE_OK', { exact: true }).waitFor();
  };
  await check('index.html boots and connects', async () => { await page.goto(`${url}/index.html`); await connected(); });
  await check('header icons share geometry and alignment', async () => {
    const icons = await page.locator('.header-actions .icon-button').evaluateAll(buttons => buttons.map(button => {
      const box = button.getBoundingClientRect(), svg = button.querySelector('svg')?.getBoundingClientRect();
      return { width: box.width, height: box.height, svgWidth: svg?.width, svgHeight: svg?.height,
        dx: svg ? svg.x + svg.width / 2 - box.x - box.width / 2 : null,
        dy: svg ? svg.y + svg.height / 2 - box.y - box.height / 2 : null };
    }));
    assert.equal(icons.length, 4);
    for (const icon of icons) assert.deepEqual(icon, { width: 30, height: 30, svgWidth: 20, svgHeight: 20, dx: 0, dy: 0 });
    await page.locator('.header-actions').screenshot({ path: path.join(output, 'header-icons.png') });
  });
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
    await page.keyboard.press('Escape');
  });
  await check('add, test and enable an external MCP server', async () => {
    await settings();
    assert.equal(await page.locator('.mcp-copy-binding').count(), 0);
    assert.equal(await page.getByText('Gateway access', { exact: true }).count(), 0);
    await page.locator('.mcp-settings summary').click();
    await page.locator('.mcp-form [name="name"]').fill('Fixture MCP');
    await page.locator('.mcp-form [name="target"]').fill(process.execPath);
    await page.locator('.mcp-form [name="args"]').fill(JSON.stringify([path.resolve('tests/support/mcp-fixture.mjs')]));
    for (const width of [1200, 320]) {
      await page.setViewportSize({ width, height: 900 });
      await page.locator('.mcp-form').scrollIntoViewIfNeeded();
      await page.screenshot({ path: path.join(output, `mcp-settings-${width}.png`), fullPage: true, animations: 'disabled' });
    }
    await page.setViewportSize({ width: 1200, height: 900 });
    const savedResponse = page.waitForResponse(response => response.url().endsWith('/session/mcp') && response.request().method() === 'POST' && response.request().postDataJSON()?.action === 'save');
    await page.locator('.mcp-form button[type="submit"]').click();
    const savedId = (await savedResponse).request().postDataJSON().config.id;
    await page.locator('.mcp-server-row').filter({ hasText: 'Fixture MCP' }).waitFor();
    await page.getByRole('button', { name: 'Test connection', exact: true }).click();
    await page.getByRole('status').filter({ hasText: '3 tools available' }).waitFor();
    await page.locator('[data-close="settings"]').click();
    await page.locator('#tools-button').click();
    await page.locator('#composer-tools').click();
    const selectedResponse = page.waitForResponse(response => response.url().endsWith('/session/mcp') && response.request().method() === 'POST' && response.request().postDataJSON()?.action === 'select');
    await page.locator('.mcp-picker label').filter({ hasText: 'Fixture MCP' }).locator('input').check();
    await selectedResponse;
    await page.keyboard.press('Escape');
    const auth = (await savedResponse).request().headers().authorization;
    const registered = await (await page.request.get(`${url}/internal/tools`, { headers: { authorization: auth } })).json();
    const echo = registered.tools.find(tool => tool.name.endsWith('__echo'));
    assert.ok(echo); assert.ok(!registered.tools.some(tool => tool.name === 'mcp_call_tool'));
    assert.equal(await page.locator('#provider-select').inputValue(), 'codex', 'MCP fixture must use the selected codex mock');
    mockCalls.get('codex').push({ tool: echo.name, args: { text: 'MCP_UI_RESULT' } });
    await page.locator('#prompt-input').fill('Use the selected MCP echo tool');
    await page.locator('#send-button').click();
    await page.getByRole('dialog', { name: 'Allow external MCP tool?' }).waitFor();
    await page.getByRole('button', { name: 'Allow', exact: true }).click();
    await page.getByText('WEB_SMOKE_OK', { exact: true }).waitFor();
    assert.equal(await page.locator('.bubble.error').count(), 0);
    mockCalls.get('codex').length = 0;
    await page.reload(); await connected();
    await page.locator('#tools-button').click(); await page.locator('#composer-tools').click();
    await page.locator('.mcp-picker label').filter({ hasText: 'Fixture MCP' }).locator('input').waitFor();
    assert.equal(await page.locator('.mcp-picker label').filter({ hasText: 'Fixture MCP' }).locator('input').isChecked(), true);
    await page.keyboard.press('Escape');
  });
  await check('History and Changes remain compact popovers', async () => {
    for (const name of ['history', 'changes']) {
      await page.locator(`#${name}-button`).click();
      assert.equal(await page.locator(`#panel-${name}`).evaluate(el => el.matches(':popover-open')), true);
      const box = await page.locator(`#panel-${name}`).boundingBox();
      assert.ok(box.width <= 340);
      await page.keyboard.press('Escape');
    }
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
  await check('system controls use authenticated APIs and report the local MCP catalog', async () => {
    assert.equal((await page.request.get(`${url}/session/system`)).status(), 401);
    await page.locator('#system-version').scrollIntoViewIfNeeded();
    await page.locator('#system-providers button').first().waitFor();
    await page.locator('#system-check-mcp').click();
    await page.waitForFunction(() => document.querySelector('#system-report')?.textContent.includes('MCP_CATALOG_OK'));
    await page.waitForFunction(() => !document.querySelector('#system-export')?.disabled);
    const download = page.waitForEvent('download'); await page.locator('#system-export').click();
    const downloaded = await download, reportPath = path.join(output, 'system-diagnostics.json');
    await downloaded.saveAs(reportPath);
    const report = JSON.parse(await readFile(reportPath, 'utf8'));
    assert.equal(report.schemaVersion, 1); assert.equal(report.scope.mcp, true);
    assert.ok(report.checks.some(item => item.code === 'MCP_CATALOG_OK'));
    assert.ok(!JSON.stringify(report).includes(temporary));
    assert.ok(!JSON.stringify(report).includes(os.homedir()));
  });
  await check('temporary debug controls persist and stop their expiry', async () => {
    await page.locator('#system-debug-on').click();
    await page.waitForFunction(() => document.querySelector('#system-status')?.textContent.includes('enabled for 15 minutes'));
    const enabled = JSON.parse(await readFile(path.join(temporary, 'debug.json'), 'utf8'));
    assert.ok(enabled.expires > Date.now() && enabled.expires <= Date.now() + 15 * 60000);
    await page.locator('#system-debug-off').click();
    await page.waitForFunction(() => document.querySelector('#system-status')?.textContent.includes('Debug logging disabled'));
    assert.equal(JSON.parse(await readFile(path.join(temporary, 'debug.json'), 'utf8')).expires, 0);
  });
  await check('rollback explains unavailable cached releases without starting an updater', async () => {
    await page.locator('#system-rollback').click();
    await page.waitForFunction(() => document.querySelector('#system-status')?.textContent.includes('No retained version'));
  });
  await check('recovery controls fit a narrow pane in both themes', async () => {
    for (const theme of ['light', 'dark']) {
      await page.setViewportSize({ width: 320, height: 900 }); await page.locator('#theme-select').selectOption(theme);
      await page.locator('#settings-system-section').scrollIntoViewIfNeeded();
      const bounds = await page.locator('#settings-system-section').evaluate(el => ({ width: el.clientWidth, scroll: el.scrollWidth }));
      assert.ok(bounds.scroll <= bounds.width + 2, `recovery controls overflow in ${theme}: ${JSON.stringify(bounds)}`);
      await page.screenshot({ path: path.join(output, `recovery-320-${theme}.png`), fullPage: true, animations: 'disabled' });
    }
    await page.setViewportSize({ width: 1200, height: 900 });
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
          await page.screenshot({ path: path.join(output, `${panel}-${width}-${theme}.png`), fullPage: true, animations: 'disabled' });
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
