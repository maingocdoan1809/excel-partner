// Excel Partner QA harness (managed file).
// What the page looked like when a test failed: open dialogs, fields marked invalid, messages, loading state.
// It is written next to the failure screenshot so a failed step can be repaired without guessing.

export async function collectPageState(page) {
  const read = page.evaluate(() => {
    const visible = (el) => {
      const r = el.getBoundingClientRect();
      const s = getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
    };
    const squash = (value, n) => String(value || '').replace(/\s+/g, ' ').trim().slice(0, n);
    const text = (el, n) => squash(el.innerText || el.textContent, n);
    const all = (selector) => Array.from(document.querySelectorAll(selector)).filter(visible);
    return {
      url: location.pathname,
      title: squash(document.title, 80),
      dialogs: all('[role=dialog], dialog[open], .modal.show, .o_dialog').slice(0, 3).map((d) => ({
        title: text(d.querySelector('.modal-title, header, h1, h2, h3, h4') || d, 80),
        text: text(d, 240),
        buttons: Array.from(d.querySelectorAll('button, .btn')).filter(visible).map((b) => text(b, 30)).filter(Boolean).slice(0, 6),
      })),
      invalid: all('.o_field_invalid, .is-invalid, [aria-invalid=true]').slice(0, 8)
        .map((e) => squash(e.getAttribute('name') || (e.closest('[name]') && e.closest('[name]').getAttribute('name')) || e.textContent, 30)),
      notices: all('.o_notification, .alert-danger, .alert-warning, [role=alert]').slice(0, 3).map((n) => text(n, 160)),
      busy: all('.o_loading_indicator, .o_blockUI, .o_ui_blocked').length > 0,
    };
  });
  // A hung page must not hold up the report.
  const timeout = new Promise((resolve) => setTimeout(() => resolve(null), 4000));
  try { return await Promise.race([read, timeout]); } catch { return null; }
}
