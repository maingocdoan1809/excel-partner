# -*- coding: utf-8 -*-
"""READ-ONLY DOM probe: opens a route and dumps the interactive structure (headings, buttons, inputs, tabs, tables, dialogs) with the
attributes that make stable selectors (id, name, title, aria-label, data-testid, placeholder, role). Output: docs/dom/<route>.json + a short print.
  python -m qakit.tools.dom_probe /knowledge /groups ...   (never clicks anything)"""
import json, re, sys
from ..core import browser, ui
from ..core.config import cfg, WORK

JS = """()=>{
 const vis=e=>e.offsetParent!==null||getComputedStyle(e).position==='fixed';
 const at=e=>({tag:e.tagName.toLowerCase(),id:e.id||undefined,name:e.getAttribute('name')||undefined,type:e.getAttribute('type')||undefined,
   role:e.getAttribute('role')||undefined,title:e.getAttribute('title')||undefined,aria:e.getAttribute('aria-label')||undefined,
   testid:e.getAttribute('data-testid')||undefined,ph:e.getAttribute('placeholder')||undefined,href:e.getAttribute('href')||undefined,
   text:(e.innerText||e.value||'').trim().slice(0,60)||undefined,dis:e.disabled||undefined,cls:(e.className&&e.className.toString().slice(0,50))||undefined});
 const q=s=>[...document.querySelectorAll(s)].filter(vis).map(at);
 return {url:location.pathname+location.search,title:document.title,
  headings:q('h1,h2,h3').map(x=>x.text),
  nav:q('nav a,nav button,aside a,aside button').slice(0,40),
  buttons:q('button').slice(0,120),inputs:q('input,textarea,select').slice(0,60),
  tabs:q('[role=tab]'),tables:[...document.querySelectorAll('table')].filter(vis).map(t=>({headers:[...t.querySelectorAll('th')].map(x=>x.innerText.trim()),rows:t.querySelectorAll('tbody tr').length})),
  dialogs:q('[role=dialog],[role=alertdialog]'),links:q('a[href]').slice(0,40),
  menuTriggers:q("button[aria-haspopup='menu']").length, checkboxes:q('[role=checkbox]').length}}"""


def probe(pg, route):
    route = "/" + route.lstrip("/")
    ui.goto(pg, cfg.url(route), settle=2500)
    d = pg.evaluate(JS)
    d["requested"] = route
    out = WORK / "docs" / "dom"
    out.mkdir(parents=True, exist_ok=True)
    name = re.sub(r"[^A-Za-z0-9]+", "_", route.strip("/")) or "root"
    json.dump(d, open(out / f"{name}.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return d


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    pg = browser.page()
    pg.bring_to_front()
    for r in sys.argv[1:]:
        d = probe(pg, r)
        print(f"== {r} -> {d['url']} | h: {d['headings'][:4]} | buttons {len(d['buttons'])} inputs {len(d['inputs'])} tabs {len(d['tabs'])} tables {[t['headers'] for t in d['tables']]} menus {d['menuTriggers']}")
    browser.stop()
