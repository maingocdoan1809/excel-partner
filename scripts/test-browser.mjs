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
class ProgressMock extends MockAdapter {
  async *sendTurn(input) {
    if (input.prompt === 'BROWSER_OPENCODE_PROGRESS') {
      yield { type: 'turn_started', requestId: input.requestId };
      yield { type: 'provider_progress', requestId: input.requestId, state: 'retrying', message: 'OpenCode: model service overloaded; retrying (attempt 2)…' };
      await new Promise(resolve => setTimeout(resolve, 1000));
      yield { type: 'error', requestId: input.requestId, code: 'empty_response', message: 'OpenCode ended without a response. Try again or select another model.' };
      yield { type: 'turn_finished', requestId: input.requestId, reason: 'error' };
      return;
    }
    yield* super.sendTurn(input);
  }
}
const registry = new ProviderRegistry();
const mockCalls = new Map();
for (const id of ['claude', 'copilot', 'codex', 'opencode']) {
  const toolCalls = []; mockCalls.set(id, toolCalls);
  registry.register(new ProgressMock({ id, toolCalls, reply: 'WEB_SMOKE_OK', delayMs: 150 }));
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
    assert.ok((await page.locator('#panel-references > .panel-head').boundingBox()).height <= 52);
    await page.locator('#panel-references [data-close]').click();
    assert.equal(await page.locator('#settings-instructions').inputValue(), 'An unsaved draft');
    assert.ok(Math.abs(await page.locator('.settings-body').evaluate(el => el.scrollTop) - scroll) <= 2);
    for (const child of ['skills', 'templates']) {
      await page.locator(`#menu-${child}`).click();
      assert.ok((await page.locator(`#panel-${child} > .panel-head`).boundingBox()).height <= 52, `${child} header should stay compact`);
      await page.locator(`#panel-${child} [data-close]`).click();
    }
  });
  await check('system controls use authenticated APIs and report the local MCP catalog', async () => {
    assert.equal((await page.request.get(`${url}/session/system`)).status(), 401);
    await page.locator('#system-version').scrollIntoViewIfNeeded();
    await page.locator('#system-providers button').first().waitFor();
    assert.equal(await page.locator('#system-export').isHidden(), true, 'export appears only after a check');
    await page.locator('#system-check').click();
    await page.waitForFunction(() => document.querySelector('#system-report')?.textContent.includes('MCP_CATALOG_OK'));
    await page.waitForFunction(() => !document.querySelector('#system-export')?.hidden && !document.querySelector('#system-export')?.disabled);
    const download = page.waitForEvent('download'); await page.locator('#system-export').click();
    const downloaded = await download, reportPath = path.join(output, 'system-diagnostics.json');
    await downloaded.saveAs(reportPath);
    const report = JSON.parse(await readFile(reportPath, 'utf8'));
    assert.equal(report.schemaVersion, 1); assert.equal(report.scope.mcp, true);
    assert.ok(report.checks.some(item => item.code === 'MCP_CATALOG_OK'));
    assert.ok(!JSON.stringify(report).includes(temporary));
    assert.ok(!JSON.stringify(report).includes(os.homedir()));
  });
  await check('SQLite history exposes storage, one full JSON export and confirmed clean-up', async () => {
    await page.locator('#history-storage-section').scrollIntoViewIfNeeded();
    assert.ok((await readFile(path.join(temporary, 'history.db'))).subarray(0, 15).toString() === 'SQLite format 3');
    assert.equal(await page.locator('#history-recover-pending').isHidden(), true, 'recovery appears only while chats are unsaved');
    const download = page.waitForEvent('download'); await page.locator('#history-export').click();
    const exported = await download; const file = path.join(output, 'history-export.json'); await exported.saveAs(file);
    const data = JSON.parse(await readFile(file, 'utf8')); assert.ok(data.sessions.length >= 2);
    assert.ok(data.sessions.some(session => session.messages.some(message => message.text.includes('Browser review message'))));
    await page.waitForFunction(() => document.querySelector('#history-storage-status')?.textContent.includes('History exported'));
  });
  await check('temporary debug controls persist and stop their expiry', async () => {
    await page.locator('#system-debug').click();
    await page.waitForFunction(() => document.querySelector('#system-status')?.textContent.includes('enabled for 15 minutes'));
    const enabled = JSON.parse(await readFile(path.join(temporary, 'debug.json'), 'utf8'));
    assert.ok(enabled.expires > Date.now() && enabled.expires <= Date.now() + 15 * 60000);
    await page.waitForFunction(() => document.querySelector('#system-debug')?.textContent === 'Stop debug');
    await page.locator('#system-debug').click();
    await page.waitForFunction(() => document.querySelector('#system-status')?.textContent.includes('Debug logging disabled'));
    assert.equal(JSON.parse(await readFile(path.join(temporary, 'debug.json'), 'utf8')).expires, 0);
  });
  await check('there is no rollback control: older releases are not kept', async () => {
    assert.equal(await page.locator('#system-rollback, #system-rollback-version').count(), 0);
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
  await check('Settings actions and fields stay aligned and readable in small panels', async () => {
    for (const width of [280, 320, 380]) for (const theme of ['light', 'dark']) {
      await page.setViewportSize({ width, height: 720 }); await page.locator('#theme-select').selectOption(theme);
      assert.ok((await page.locator('#panel-settings > .panel-head').boundingBox()).height <= 52, `Settings header should stay compact at ${width}`);
      for (const section of ['settings-system-section', 'history-storage-section']) {
        await page.locator(`#${section}`).scrollIntoViewIfNeeded();
        const layout = await page.locator(`#${section}`).evaluate(el => {
          const bounds = el.getBoundingClientRect();
          return { width: el.clientWidth, scroll: el.scrollWidth, buttons: [...el.querySelectorAll('button')].filter(b => !b.hidden).map(b => {
            const r = b.getBoundingClientRect(); return { left: r.left, right: r.right, height: r.height, width: r.width, scroll: b.scrollWidth, client: b.clientWidth, inside: r.left >= bounds.left && r.right <= bounds.right };
          }) };
        });
        assert.ok(layout.scroll <= layout.width + 1, `${section} must not overflow at ${width}`);
        assert.ok(layout.buttons.every(b => b.inside && b.height >= 36 && b.scroll <= b.client + 1), `buttons must fit readable touch targets at ${width}: ${JSON.stringify(layout)}`);
        if (width <= 320) {
          const ids = section === 'settings-system-section' ? ['system-check', 'system-export', 'system-repair', 'system-debug']
            : ['history-export', 'settings-clear-history'];
          const actions = await page.locator(ids.map(id => `#${id}`).join(',')).evaluateAll(buttons => buttons.map(b => { const r = b.getBoundingClientRect(); return { left: r.left, right: r.right }; }));
          assert.ok(actions.every(b => Math.abs(b.left - actions[0].left) <= 1 && Math.abs(b.right - actions[0].right) <= 1), `small-panel actions must align at ${width}`);
        }
        await page.screenshot({ path: path.join(output, `${section}-${width}-${theme}.png`), animations: 'disabled' });
      }
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
          const header = await page.locator(`#panel-${panel} > .panel-head`).boundingBox();
          assert.ok(header.height >= 40 && header.height <= 52, `${panel} header should stay compact at ${width}: ${header.height}px`);
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
  await check('pending history survives closing a pane and is saved when a new pane reconnects', async () => {
    const context = await browser.newContext();
    let tab = await context.newPage();
    try {
      await tab.route('**/session/history', route => {
        const request = route.request();
        if (request.method() === 'POST' && request.postDataJSON()?.action === 'save') return route.fulfill({ status: 503, contentType: 'application/json', body: JSON.stringify({ error: 'Fixture history storage unavailable', code: 'HISTORY_STORAGE_FAILED' }) });
        return route.continue();
      });
      await tab.goto(`${url}/index.html`); await tab.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
      await tab.locator('#prompt-input').fill('PENDING_CLOSE_RECOVERY'); await tab.locator('#send-button').click();
      await tab.getByText('WEB_SMOKE_OK', { exact: true }).waitFor();
      await tab.waitForFunction(() => /could not be saved|not saved/i.test(document.querySelector('#history-save-status')?.textContent ?? ''));
      assert.ok(await tab.evaluate(() => Object.keys(localStorage).some(key => key.startsWith('excelPartner.historyOutbox.v1.') && localStorage.getItem(key).includes('PENDING_CLOSE_RECOVERY'))));
      await tab.close(); tab = await context.newPage(); await tab.goto(`${url}/index.html`);
      await tab.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
      await tab.locator('#history-button').click(); await tab.locator('.history-item').filter({ hasText: 'PENDING_CLOSE_RECOVERY' }).waitFor();
      assert.equal(await tab.locator('.history-item').filter({ hasText: 'PENDING_CLOSE_RECOVERY' }).count(), 1);
    } finally { await context.close(); }
  });
  await check('legacy history migrates once and retains the original source for rollback', async () => {
    const legacy = await browser.newPage();
    try {
      await legacy.goto(`${url}/index.html`);
      await legacy.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
      await legacy.evaluate(() => {
        const now = Date.now();
        localStorage.setItem('excelPartner.sessions.v1', JSON.stringify(['excel', 'word'].map(host => ({ id: 'legacy-' + host, provider: 'claude', host, title: 'Legacy ' + host,
          startedAt: now, updatedAt: now, messages: [{ kind: 'user', text: 'LEGACY_' + host.toUpperCase() }], changes: [] }))));
        localStorage.removeItem('excelPartner.historyMigrated.v1.excel');
      });
      await legacy.reload();
      await legacy.waitForFunction(() => localStorage.getItem('excelPartner.historyMigrated.v1.excel') === '1');
      await legacy.locator('#history-button').click();
      await legacy.locator('.history-item').filter({ hasText: 'Legacy excel' }).waitFor();
      assert.equal(await legacy.locator('.history-item').filter({ hasText: 'Legacy word' }).count(), 0);
      assert.equal(await legacy.evaluate(() => JSON.parse(localStorage.getItem('excelPartner.sessions.v1')).length), 2);
      await legacy.reload(); await legacy.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
      await legacy.locator('#history-button').click(); await legacy.locator('.history-item').filter({ hasText: 'Legacy excel' }).waitFor();
      assert.equal(await legacy.locator('.history-item').filter({ hasText: 'Legacy excel' }).count(), 1);
    } finally { await legacy.close(); }
  });
  await check('two panes discard stale preexisting outbox copies without blocking History or losing saved messages', async () => {
    const context = await browser.newContext(); let tab = await context.newPage(), token;
    tab.on('websocket', socket => socket.on('framereceived', frame => { try { const value = JSON.parse(String(frame.payload)); if (value.type === 'hello_ok') token = value.sessionToken; } catch {} }));
    try {
      await tab.goto(`${url}/index.html`); await tab.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
      assert.ok(token);
      await tab.evaluate(async token => {
        const headers = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };
        const stats = await (await fetch('/session/history?view=stats', { headers })).json();
        const old = { id: 'existing-history-sync', host: 'excel', provider: 'copilot', title: 'EXISTING_HISTORY_SYNC', startedAt: 1, updatedAt: 2, messages: [{ kind: 'user', text: 'EXISTING_HISTORY_SYNC' }], changes: [] };
        const newer = { ...old, messages: [...old.messages, { kind: 'assistant', text: 'Already committed answer' }] };
        const response = await fetch('/session/history', { method: 'POST', headers, body: JSON.stringify({ action: 'save', session: newer, baseRevision: 0, generation: stats.generation, operationId: 'existing-sync-server' }) });
        if (!response.ok) throw Error('Could not seed the saved fixture');
        localStorage.setItem('excelPartner.sessions.v1', JSON.stringify([old])); localStorage.removeItem('excelPartner.historyMigrated.v1.excel');
        localStorage.setItem('excelPartner.historyOutbox.v1.closed-pane', JSON.stringify([{ session: old, host: 'excel', baseRevision: 0, generation: stats.generation, operationId: 'delayed-existing-snapshot' }]));
      }, token);
      const second = await context.newPage(); await Promise.all([tab.reload(), second.goto(`${url}/index.html`)]);
      for (const pane of [tab, second]) {
        await pane.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
        await pane.locator('#history-button').click(); const row = pane.locator('.history-item').filter({ hasText: 'EXISTING_HISTORY_SYNC' });
        await row.waitFor(); assert.equal(await row.count(), 1); await row.click();
        await pane.getByText('Already committed answer', { exact: true }).waitFor();
        assert.ok(!(await pane.locator('#history-save-status').innerText()).includes('History is unavailable'));
      }
      await tab.waitForFunction(() => !Object.keys(localStorage).some(key => key.startsWith('excelPartner.historyOutbox.v1.') && (localStorage.getItem(key) ?? '').includes('delayed-existing-snapshot')));
    } finally { await context.close(); }
  });
  await check('OpenCode retries are visible and an empty response releases the composer', async () => {
    const context = await browser.newContext(), pane = await context.newPage();
    try {
      await pane.goto(`${url}/index.html`);
      await pane.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
      await pane.locator('#menu-button').click(); await pane.locator('#provider-select').selectOption('opencode');
      await pane.locator('[data-close="settings"]').click();
      await pane.locator('#prompt-input').fill('BROWSER_OPENCODE_PROGRESS'); await pane.locator('#send-button').click();
      await pane.locator('.thinking-label').filter({ hasText: 'model service overloaded; retrying' }).waitFor();
      await pane.getByText('OpenCode ended without a response. Try again or select another model.', { exact: true }).waitFor();
      assert.equal(await pane.locator('.thinking-label').count(), 0);
      await pane.waitForFunction(() => document.querySelector('#send-button')?.dataset.mode !== 'stop');
      assert.equal(await pane.locator('#prompt-input').isEditable(), true);
    } finally { await context.close(); }
  });
  await check('root route also boots' , async () => { await page.goto(url); await connected(); });
  await check('no uncaught browser errors or failed requests', async () => { assert.deepEqual(errors, []); assert.deepEqual(failedRequests, []); });
} finally {
  await browser?.close();
  await bridge.close();
  await rm(temporary, { recursive: true, force: true });
  await writeFile(path.join(output, 'results.json'), JSON.stringify({ checks, errors, failedRequests }, null, 2));
}
if (checks.some(check => !check.ok)) process.exitCode = 1;
console.log(`Browser artifacts: ${output}`);
