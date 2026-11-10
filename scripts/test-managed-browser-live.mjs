/** Opt-in real browser + managed Copilot + a harmless local MCP echo tool. Costs model tokens. */
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';
import { Bridge } from '../dist/bridge/src/server.js';
import { loadConfig } from '../dist/bridge/src/config.js';
import store from '../bin/runtime-store.cjs';

if (process.env.LIVE_MANAGED_BROWSER !== 'copilot') throw new Error('Opt in with LIVE_MANAGED_BROWSER=copilot.');
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'output/playwright/managed-browser-live'); fs.mkdirSync(output, { recursive: true });
const runtimeCheck = JSON.parse(fs.readFileSync(path.join(root, 'output/playwright/runtime-live/results.json'), 'utf8'));
const dataDir = path.resolve(root, runtimeCheck.runtimeData);
if (!store.inside(path.join(root, 'output/playwright/runtime-live'), dataDir)) throw new Error('Invalid opt-in runtime fixture directory.');
const profile = store.selectProfile(dataDir, 'copilot');
if (profile?.mode !== 'managed') throw new Error('Run the opt-in managed Copilot installation check first.');
process.env.OFFICE_PARTNER_DATA_DIR = dataDir;
process.env.BRIDGE_CONFIG = path.join(dataDir, 'config.json');
// This directory belongs to this opt-in fixture; keep reruns independent of earlier MCP selections.
fs.writeFileSync(path.join(dataDir, 'mcp-servers.json'), '[]');
const port = Number(process.env.BROWSER_TEST_PORT ?? 8765), certPath = process.env.BROWSER_TEST_CERT, keyPath = process.env.BROWSER_TEST_KEY;
if (!certPath || !keyPath) throw new Error('Provide BROWSER_TEST_CERT and BROWSER_TEST_KEY for local HTTPS.');
const url = `https://localhost:${port}`;
const config = loadConfig({ host: '127.0.0.1', port, tls: { certPath, keyPath }, publicBaseUrl: url, configFile: process.env.BRIDGE_CONFIG,
  skillDirs: [], referenceDirs: [], logLevel: 'error' });
const bridge = new Bridge(config, { usageFile: null, updater: null, workspaceConfigPath: null,
  skillUploadDir: path.join(dataDir, 'skills'), systemPromptFile: path.join(dataDir, 'prompt.txt'), savedReferencesFile: path.join(dataDir, 'references.json') });
let browser;
const checks = [], errors = [], tools = [];
try {
  await bridge.listen();
  browser = await chromium.launch({ headless: false, ...(process.env.BROWSER_EXECUTABLE ? { executablePath: process.env.BROWSER_EXECUTABLE } : { channel: 'msedge' }) });
  const page = await browser.newPage({ ignoreHTTPSErrors: true, viewport: { width: 1100, height: 850 } });
  page.setDefaultTimeout(120000); page.on('pageerror', error => errors.push(error.message));
  page.on('websocket', socket => socket.on('framereceived', frame => {
    try { const event = JSON.parse(String(frame.payload)); if (event.type === 'tool_finished') tools.push({ tool: event.tool, ok: event.ok }); } catch { /* non-JSON frame */ }
  }));
  await page.goto(`${url}/index.html`);
  await page.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
  await page.locator('#menu-button').click(); await page.locator('#provider-select').selectOption('copilot');
  await page.locator('.mcp-settings summary').click();
  await page.locator('.mcp-form [name="name"]').fill('Managed runtime echo');
  await page.locator('.mcp-form [name="target"]').fill(process.execPath);
  await page.locator('.mcp-form [name="args"]').fill(JSON.stringify([path.join(root, 'tests/support/mcp-fixture.mjs')]));
  await page.locator('.mcp-form button[type="submit"]').click();
  await page.locator('.mcp-server-row').filter({ hasText: 'Managed runtime echo' }).waitFor();
  await page.locator('[data-close="settings"]').click(); await page.locator('#tools-button').click(); await page.locator('#composer-tools').click();
  await page.locator('.mcp-picker label').filter({ hasText: 'Managed runtime echo' }).locator('input').check(); await page.keyboard.press('Escape');
  async function turn(marker) {
    await page.locator('#prompt-input').fill(`Use the enabled MCP tool whose name ends in __echo exactly once with text ${marker}. Do not use other tools. Then reply exactly ${marker}.`);
    await page.locator('#send-button').click();
    const dialog = page.getByRole('dialog', { name: 'Allow external MCP tool?' }); await dialog.waitFor();
    await page.getByRole('button', { name: 'Allow', exact: true }).click();
    await page.waitForFunction(marker => [...document.querySelectorAll('.bubble.assistant')].some(el => el.textContent.includes(marker)), marker);
    await page.waitForFunction(() => !document.querySelector('#prompt-input')?.disabled && document.querySelector('#stop-button')?.hidden !== false);
    assert.equal(await page.locator('.bubble.error').count(), 0);
  }
  await turn('PRIVATE_CLI_BROWSER_OK');
  assert.ok(tools.some(tool => tool.tool.endsWith('__echo') && tool.ok));
  checks.push({ name: 'Real managed CLI invokes native MCP schema through browser approval', ok: true });
  console.log('PASS real managed Copilot + native external MCP + browser approval.');
  await turn('PRIVATE_CLI_CONTINUATION_OK');
  assert.equal(tools.filter(tool => tool.tool.endsWith('__echo') && tool.ok).length, 2);
  checks.push({ name: 'Second turn preserves the native conversation and executes once', ok: true });
  console.log('PASS same native conversation, second tool turn, no replay.');
  await page.reload();
  await page.waitForFunction(() => document.querySelector('#connection-status')?.dataset.state === 'open');
  await page.locator('#history-button').click(); await page.locator('.history-item').first().click();
  await turn('PRIVATE_CLI_RESUME_OK');
  assert.equal(tools.filter(tool => tool.tool.endsWith('__echo') && tool.ok).length, 3);
  assert.ok(Object.values(store.readState(dataDir).conversations).includes(profile.id));
  checks.push({ name: 'Reload and History resume retain the managed native profile without replay', ok: true });
  console.log('PASS reload + History resume with the recorded private runtime.');
  const after = store.selectProfile(dataDir, 'copilot'); assert.equal(after.id, profile.id);
  assert.deepEqual(errors, []);
  await page.screenshot({ path: path.join(output, 'copilot-private-cli-mcp.png'), fullPage: true });
} catch (error) {
  checks.push({ name: 'Real managed CLI browser flow', ok: false, error: error.message });
  if (browser) await browser.contexts()[0]?.pages()[0]?.screenshot({ path: path.join(output, 'failure.png'), fullPage: true }).catch(() => {});
  console.error(error.message); process.exitCode = 1;
} finally {
  await browser?.close(); await bridge.close();
  fs.writeFileSync(path.join(output, 'results.json'), JSON.stringify({ provider: 'copilot', version: profile.version, checks, tools, errors }, null, 2));
}
