"""WHY did the screen not do what the step expected? Read what the page itself says instead of guessing. Nothing here raises: a failed
diagnosis returns a summary that says so."""

_JS = """
() => {
  const vis = e => { const b = e.getBoundingClientRect(), s = getComputedStyle(e); return b.width > 1 && b.height > 1 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const txt = e => (e.innerText || e.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 240);
  const all = (sels) => sels.flatMap(s => [...document.querySelectorAll(s)]).filter(vis);
  const errors = [...new Set(all(['.alert-danger', '.o_notification.border-danger', '.o_error_dialog', '.o_form_error_dialog', '[role=alert]', '.invalid-feedback']).map(txt).filter(Boolean))];
  const messages = [...new Set(all(['.o_notification', '.alert-warning', '.modal .modal-body']).map(txt).filter(Boolean))];
  const label = e => { const w = e.closest('.o_field_widget, .o_wrap_field, .mb-3, td, div'); const l = w && (w.querySelector('label') || w.previousElementSibling);
    return (l && txt(l)) || e.getAttribute('name') || e.getAttribute('placeholder') || e.tagName.toLowerCase(); };
  const fields = [], seen = new Set();
  const add = (e, why) => { const i = e.matches('input,select,textarea') ? e : (e.querySelector('input,select,textarea') || e); const l = label(i);
    if (!seen.has(l)) { seen.add(l); fields.push({label: l, why, name: (e.closest('[name]') || i).getAttribute('name') || ''}); } };
  for (const e of all(['.o_field_invalid', '[aria-invalid=true]', '.is-invalid'])) add(e, 'flagged invalid');
  for (const e of all(['.o_required_modifier'])) { const i = e.matches('input,select,textarea') ? e : e.querySelector('input,select,textarea');
    if (i && !i.disabled && !i.readOnly && !String(i.value || '').trim()) add(e, 'required and empty'); }
  const dialogs = all(['[role=dialog]', '.modal.show', '.o_dialog']).map(txt);
  return {errors: errors.slice(0, 5), messages: messages.slice(0, 5), fields: fields.slice(0, 12), dialogs: dialogs.slice(0, 2), url: location.href};
}
"""


def read(page):
    out = {"summary": "", "errors": [], "messages": [], "fields": [], "dialogs": [], "url": ""}
    try:
        out.update(page.evaluate(_JS))
        parts = []
        if out["fields"]:
            parts.append("fields: " + "; ".join(f"{f['label']} ({f['why']})" for f in out["fields"]))
        if out["errors"]:
            parts.append("errors: " + " | ".join(out["errors"][:3]))
        elif out["messages"]:
            parts.append("messages: " + " | ".join(out["messages"][:3]))
        if out["dialogs"]:
            parts.append("dialog: " + out["dialogs"][0][:120])
        out["summary"] = " ; ".join(parts) or "the page shows no error message or invalid field (look at the picture)"
    except Exception as error:
        out["summary"] = f"diagnosis failed: {str(error)[:120]}"
    return out
