// Excel Partner QA harness (managed file).
// Signs in once and saves the browser session. Run through `excel-partner qa login`, not directly.
//   interactive: a visible browser opens and the person types the password (and any code).
//   headless:    credentials arrive as JSON on stdin and are typed by this script.
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from '@playwright/test';
import { launchArgs } from './launch.js';

const env = process.env;
const mode = env.QA_LOGIN_MODE === 'headless' ? 'headless' : 'interactive';
const out = env.QA_STORAGE_OUT;
const loginTarget = new URL(env.QA_LOGIN_PATH || '/web/login', env.QA_BASE_URL);
if (loginTarget.origin !== new URL(env.QA_BASE_URL).origin) {
  console.error('Sign-in did not start: the login path leaves the environment origin.');
  process.exit(1);
}
const success = new RegExp(env.QA_LOGIN_SUCCESS_URL || '/(odoo|web)(?!/login)');
const timeoutMs = Number(env.QA_LOGIN_TIMEOUT_MS || 300000);
const odooCheck = env.QA_CHECK_TYPE === 'odoo-session';

async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(chunk);
  return Buffer.concat(chunks).toString('utf8');
}

const credentials = mode === 'headless' ? JSON.parse(await readStdin()) : null;
const browser = await chromium.launch({ headless: mode === 'headless', args: launchArgs() });

/** Does the server accept this browser session? Asked from inside the page, so it stays within the allowlist. */
async function serverAccepts(page) {
  try {
    return await page.evaluate(async (checkPath) => {
      const response = await window.fetch(checkPath, {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'same-origin',
        body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: {} }),
      });
      const body = await response.json();
      return Boolean(body && body.result && body.result.uid);
    }, env.QA_CHECK_PATH || '/web/session/get_session_info');
  } catch { return false; } // mid-navigation, or the answer was not JSON
}

try {
  const context = await browser.newContext({ ignoreHTTPSErrors: env.QA_IGNORE_HTTPS_ERRORS === '1', viewport: { width: 1280, height: 860 } });
  const page = await context.newPage();
  await page.goto(loginTarget.toString());
  if (mode === 'headless') {
    await page.locator(env.QA_USER_SELECTOR || "input[name='login']").fill(credentials.user);
    await page.locator(env.QA_PASSWORD_SELECTOR || "input[name='password']").fill(credentials.password);
    await page.locator(env.QA_SUBMIT_SELECTOR || "button[type='submit']").first().click();
  } else {
    console.error('A browser window is open. Sign in there (including any verification code); it closes by itself.');
  }

  // Signed in = the server says so (Odoo), or the address left the login page (other sites).
  // Never wait for "network idle": Odoo keeps a long-poll request open for as long as the page is open.
  const started = Date.now();
  let lastUrl = '', lastNote = Date.now();
  for (;;) {
    if (page.isClosed()) {
      throw new Error(`the browser was closed before sign-in was detected (last page: ${lastUrl || 'unknown'}). If you had signed in, check that the address is covered by login.successUrl in qa.config.json.`);
    }
    lastUrl = page.url();
    const signedIn = odooCheck ? await serverAccepts(page) : (success.test(lastUrl) && !lastUrl.includes('/login'));
    if (signedIn) break;
    if (mode === 'headless') {
      const alert = page.locator('.alert-danger, .alert-warning, .alert').first();
      if (lastUrl.includes('/login') && Date.now() - started > 1500 && await alert.isVisible().catch(() => false)) {
        throw new Error(`the site rejected the sign-in: ${((await alert.innerText().catch(() => '')) || 'no reason shown').replace(/\s+/g, ' ').slice(0, 120)}`);
      }
    } else if (Date.now() - lastNote > 20000) {
      console.error(`Still waiting for sign-in to be accepted... current page: ${lastUrl}`);
      lastNote = Date.now();
    }
    if (Date.now() - started > timeoutMs) throw new Error(`timed out waiting for sign-in (last page: ${lastUrl})`);
    await page.waitForTimeout(700).catch(() => {});
  }
  await page.waitForLoadState('domcontentloaded').catch(() => {});
  fs.mkdirSync(path.dirname(out), { recursive: true });
  await context.storageState({ path: out });
  try { fs.chmodSync(out, 0o600); } catch { /* not supported on every file system */ }
  console.log('Session saved.');
} catch (error) {
  console.error(`Sign-in did not complete: ${String(error.message || error).split('\n')[0]}`);
  process.exitCode = 1;
} finally {
  await browser.close().catch(() => {});
}
