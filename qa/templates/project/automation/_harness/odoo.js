// Excel Partner QA harness (managed file).
// Helpers for Odoo web clients (14-18). They only use the page, so every request stays inside the origin allowlist.
import { expect } from '@playwright/test';

/** Models that could escalate privileges or run server-side code; blocked from odoo.rpc. */
const DENIED_MODELS = new Set([
  'res.users', 'res.groups', 'ir.config_parameter', 'ir.rule', 'ir.model.access', 'ir.cron', 'ir.actions.server',
  'ir.module.module', 'base.module.upgrade', 'ir.model', 'ir.model.fields', 'ir.ui.view', 'ir.mail_server', 'fetchmail.server',
]);
const DENIED_METHODS = new Set(['sudo', 'with_user', 'with_env', 'execute', 'execute_kw', 'init', 'update_list', 'button_immediate_install', 'button_immediate_upgrade']);

export function odooHelpers(page) {
  const helpers = {
    /** Wait until the web client is rendered and no blocking spinner is shown. */
    async ready() {
      await page.waitForSelector('.o_web_client', { state: 'attached' });
      await page.waitForFunction(() => !document.querySelector('.o_loading_indicator, .o_blockUI, .o_ui_blocked'), null, { timeout: 30000 }).catch(() => {});
      // Not "network idle": Odoo keeps a long-poll request open, so the network never goes idle.
      await page.waitForSelector('.o_action_manager, .o_home_menu, .o_main_navbar', { state: 'attached', timeout: 30000 }).catch(() => {});
    },
    async goto(path = '/web') { await page.goto(path); await helpers.ready(); },
    /** Open an app from the home menu or the app switcher by its visible name. */
    async openApp(name) {
      const switcher = page.locator('.o_navbar_apps_menu button, .o_menu_toggle, .o_grid_apps_menu__button').first();
      if (await switcher.isVisible().catch(() => false)) await switcher.click();
      await page.locator('.o_app, .o-dropdown-item.o_app, .dropdown-item.o_app', { hasText: name }).first().click();
      await helpers.ready();
    },
    /** Click menu entries in the top bar, for example openMenu('Orders', 'Quotations'). */
    async openMenu(...parts) {
      for (const part of parts) {
        await page.locator('.o_menu_sections, .o_navbar, .o-dropdown--menu').locator('a, button, .dropdown-item', { hasText: part }).first().click();
        await page.waitForTimeout(150);
      }
      await helpers.ready();
    },
    async button(text) {
      await page.locator('button:visible, .btn:visible, a.btn:visible', { hasText: text }).first().click();
      await helpers.ready();
    },
    field(name) { return page.locator(`.o_field_widget[name="${name}"], .o_field_cell[name="${name}"]`).first(); },
    async fill(name, value) { await helpers.field(name).locator('input, textarea').first().fill(String(value)); },
    async save() { await page.locator('.o_form_button_save').first().click(); await helpers.ready(); },
    notification(text) { return page.locator('.o_notification', { hasText: text }); },
    async expectNotification(text) { await expect(helpers.notification(text).first()).toBeVisible(); },
    async expectNoError() { await expect(page.locator('.o_error_dialog, .modal-dialog .o_dialog_error')).toHaveCount(0); },
    /** Call an ORM method from the signed-in session (test data set-up and clean-up). */
    async rpc(model, method, args = [], kwargs = {}) {
      if (DENIED_MODELS.has(model)) throw new Error(`odoo.rpc: model ${model} is blocked in tests.`);
      if (DENIED_METHODS.has(method) || method.startsWith('_')) throw new Error(`odoo.rpc: method ${method} is blocked in tests.`);
      return page.evaluate(async ({ model, method, args, kwargs }) => {
        const response = await window.fetch(`/web/dataset/call_kw/${model}/${method}`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'same-origin',
          body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: { model, method, args, kwargs } }),
        });
        const body = await response.json();
        if (body.error) throw new Error((body.error.data && body.error.data.message) || body.error.message || 'RPC error');
        return body.result;
      }, { model, method, args, kwargs });
    },
  };
  return helpers;
}
