# -*- coding: utf-8 -*-
"""WHY did the screen not do what the step expected? Read it instead of guessing, then repair what is OURS to repair.

read(pg)                      what the page itself says is wrong, no side effects:
                                fields   invalid or required-and-empty inputs, each {label, why, value, sel} (sel = a selector that finds exactly that input)
                                errors   error banners / toasts / inline feedback (a SUCCESS toast is not an error)
                                messages every visible notification text, errors included
                                dialogs  open dialogs
snapshot(pg, label, ...)      read() + a screenshot saved to <out>/diagnose/ + a 'diagnosis' event the dashboard shows
autofill(pg, d)               fill the fields reported EMPTY with plausible valid test values (text with the test prefix, first option of a select,
                              first suggestion of an autocomplete, 1 for a number, today for a date); returns what it filled
Extra selectors for the project's UI framework: qa.config.json diagnose.selectors {"errors": [...], "invalid": [...], "required": [...]}. Odoo is covered.
Nothing here raises: a failed diagnosis returns {"summary": "diagnosis failed: ..."}."""
import datetime, re
from . import events
from .config import cfg

ERROR_SELECTORS = [".alert-danger", ".invalid-feedback", ".o_notification.border-danger", ".o_form_error_dialog", "[role=alert]", ".error", ".o_error_dialog"]
MESSAGE_SELECTORS = [".alert-warning", ".o_notification_content", ".modal .modal-body"]
INVALID_SELECTORS = [".o_field_invalid", "[aria-invalid=true]", ".is-invalid", ".has-error"]
REQUIRED_SELECTORS = ["input[required]", "select[required]", "textarea[required]", "[aria-required=true]", ".o_required_modifier"]

_JS = """
(cfg) => {
  const vis = e => { const b = e.getBoundingClientRect(); const s = getComputedStyle(e); return b.width > 1 && b.height > 1 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const txt = e => (e.innerText || e.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 240);
  const inputOf = e => e.matches('input,select,textarea') ? e : (e.querySelector('input:not([type=hidden]),select,textarea,[contenteditable=true]') || e);
  const labelOf = e => {
    const id = e.id && document.querySelector('label[for="' + CSS.escape(e.id) + '"]');
    if (id) return txt(id);
    const w = e.closest('.o_field_widget, .o_wrap_field, .form-group, .mb-3, td, div');
    const l = w && (w.querySelector('label') || (w.previousElementSibling && w.previousElementSibling.querySelector && w.previousElementSibling.querySelector('label')) || w.previousElementSibling);
    return (l && txt(l)) || e.getAttribute('name') || e.getAttribute('placeholder') || e.getAttribute('aria-label') || e.tagName.toLowerCase();
  };
  const valueOf = e => (e.value !== undefined ? e.value : txt(e)) || '';
  const errors = [], messages = [], fields = [], dialogs = [];
  const seen = new Set();
  const collect = (sels, into) => { for (const sel of sels) for (const e of document.querySelectorAll(sel)) { if (!vis(e)) continue; const t = txt(e); if (t && !seen.has(t)) { seen.add(t); into.push(t); } } };
  collect(cfg.errors, errors); collect(cfg.messages, messages);
  let n = 0;
  const addField = (e, why) => {
    const input = inputOf(e); const l = labelOf(input);
    if (fields.some(f => f.label === l)) return;
    input.setAttribute('data-qa-f', String(n));
    fields.push({label: l, why, value: valueOf(input).slice(0, 60), sel: '[data-qa-f="' + (n++) + '"]'});
  };
  for (const sel of cfg.invalid) for (const e of document.querySelectorAll(sel)) { if (vis(e)) addField(e, e.getAttribute('aria-invalid') ? 'aria-invalid' : 'flagged invalid'); }
  for (const e of document.querySelectorAll('input:invalid, select:invalid, textarea:invalid')) { if (vis(e)) addField(e, e.validationMessage || 'invalid'); }
  for (const sel of cfg.required) for (const e of document.querySelectorAll(sel)) { if (!vis(e)) continue; const i = inputOf(e); if (i.disabled || i.readOnly) continue; if (!String(valueOf(i)).trim()) addField(e, 'required and empty'); }
  for (const e of document.querySelectorAll('[role=dialog], .modal.show, .o_dialog')) { if (vis(e)) dialogs.push(txt(e)); }
  return {errors: errors.slice(0, 6), messages: messages.slice(0, 6).concat(errors).slice(0, 8), fields: fields.slice(0, 15), dialogs: dialogs.slice(0, 3), url: location.href, title: document.title};
}
"""


def _cfg():
    extra = cfg.raw.get("diagnose", {}).get("selectors", {})
    return {"errors": ERROR_SELECTORS + list(extra.get("errors", [])), "messages": MESSAGE_SELECTORS,
            "invalid": INVALID_SELECTORS + list(extra.get("invalid", [])), "required": REQUIRED_SELECTORS + list(extra.get("required", []))}


def read(pg):
    out = {"summary": "", "errors": [], "messages": [], "fields": [], "dialogs": [], "image": "", "url": ""}
    try:
        out.update(pg.evaluate(_JS, _cfg()))
        parts = []
        if out["fields"]:
            parts.append("fields: " + "; ".join(f"{f['label']} ({f['why']})" for f in out["fields"]))
        if out["errors"]:
            parts.append("errors: " + " | ".join(out["errors"][:3]))
        elif out["messages"]:
            parts.append("messages: " + " | ".join(out["messages"][:3]))
        if out["dialogs"]:
            parts.append("dialog: " + out["dialogs"][0][:120])
        out["summary"] = " ; ".join(parts) or "the page shows no validation message or invalid field (look at the picture)"
    except Exception as e:
        out["summary"] = f"diagnosis failed: {str(e)[:120]}"
    return out


def snapshot(pg, label, error=None, case=None):
    out = read(pg)
    try:
        d = cfg.out / "diagnose"
        d.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r"[^\w.-]+", "_", str(label))[:60]           # not inside the f-string: a backslash there is a syntax error before Python 3.12
        img = d / f"{datetime.datetime.now().strftime('%H%M%S')}_{safe}.png"
        pg.screenshot(path=str(img), scale="css", animations="disabled")
        out["image"] = str(img)
    except Exception:
        pass
    try:
        print(f"   DIAGNOSIS [{label}] {out['summary']}" + (f"  picture: {out['image']}" if out["image"] else ""))
        events.emit("phase", phase="fix", text=f"[{label}] {out['summary']}"[:400])
        if out["image"] and case:
            events.emit("shot", id=str(case), n=0, caption=f"diagnosis: {label}", image=out["image"])
        events.emit("diagnosis", label=str(label), error=str(error or "")[:200], url=out.get("url", ""), summary=out["summary"][:400],
                    fields=[{k: v for k, v in f.items() if k != "sel"} for f in out["fields"]], messages=out["messages"], image=out["image"], id=str(case or ""))
    except Exception:
        pass
    return out


def autofill(pg, d, prefix=None):
    """Fill the fields `d` reports EMPTY with plausible valid values. Returns ["Customer Name = ZZ_AUTOTEST_Customer Name", ...]. A field that already holds a
    value is left alone (what is wrong with it is for the case's author to decide: too long, duplicate, wrong format)."""
    prefix = prefix or cfg.prefix
    done = []
    for f in d.get("fields", []):
        sel = f.get("sel")
        if not sel or str(f.get("value", "")).strip():
            continue
        try:
            loc = pg.locator(sel).first
            tag = loc.evaluate("e => e.tagName.toLowerCase()")
            typ = (loc.get_attribute("type") or "text").lower()
            cls = (loc.get_attribute("class") or "") + " " + (loc.get_attribute("role") or "")
            label = re.sub(r"[^\w]+", "_", f["label"]).strip("_") or "value"
            if tag == "select":
                values = loc.evaluate("e => [...e.options].filter(o => o.value).map(o => o.value)")
                if not values:
                    continue
                loc.select_option(values[0]); done.append(f"{f['label']} = option {values[0]}")
            elif typ in ("checkbox", "radio", "file", "hidden"):
                continue
            elif "autocomplete" in cls or "combobox" in cls:
                loc.click()
                item = pg.locator("[role=option], .o-autocomplete--dropdown-item, .dropdown-menu .dropdown-item, .ui-menu-item").first
                item.wait_for(state="visible", timeout=3000)
                text = item.inner_text().strip()[:40]
                item.click(); done.append(f"{f['label']} = first suggestion '{text}'")
            else:
                value = {"number": "1", "date": datetime.date.today().isoformat(), "email": f"{prefix.lower()}@example.com"}.get(typ) or f"{prefix}{label}"
                ml = loc.get_attribute("maxlength")
                if ml and ml.isdigit():
                    value = value[:int(ml)]
                loc.fill(value); done.append(f"{f['label']} = {value}")
        except Exception:
            continue
    if done:
        print("   AUTOFILL " + " ; ".join(done))
        events.emit("phase", phase="fix", text="filled the empty required fields: " + " ; ".join(done)[:300])
    return done
