"""The sitemap of the Odoo under test: which menu opens which action, which views the action shows, which button leads where.

It is read from Odoo's own definitions (menus, actions, view definitions) through read-only calls: nothing is clicked, so nothing is changed and a crawl
takes minutes, not hours. A button of type `object` calls a Python method: where it leads is only known by running it, so it is listed as such.

    knowledge/sitemap.json = {schema_version, meta, nodes: {key: node}, edges: [{from, to, kind, label?, name?}], unresolved: [{where, why}]}

node keys:  menu:<id>  window:<id> | server:<id> | client:<id> | url:<id> | report:<id>  (actions)  view:<model>:<type>:<id>
edge kinds: menu (parent -> child) · opens (menu -> action) · shows (action -> view) · row (list or kanban -> form of the same model) · button (view -> action)

`build` takes any `rpc(model, method, args, kwargs)` (the signed-in `Odoo.rpc`); everything else reads the saved file, so the CLI and the MCP server share it."""
import json
import re
import xml.etree.ElementTree as ET
from collections import deque
from datetime import datetime, timezone

from .paths import workspace_dir

SCHEMA = 1
MAX_NODES = 8000
STALE_DAYS = 30
KINDS = {"ir.actions.act_window": "window", "ir.actions.server": "server", "ir.actions.client": "client", "ir.actions.act_url": "url", "ir.actions.report": "report"}
ACTION_FIELDS = {"window": ["name", "res_model", "views", "view_mode", "target"], "server": ["name", "model_name"], "client": ["name", "tag"], "url": ["name", "url"], "report": ["name", "model"]}
OPEN_ROW = ("open", "edit")                          # kanban cards: `<a type="open">` opens the record


class SitemapError(Exception):
    pass


def file_of(ws):
    return workspace_dir(ws) / "knowledge" / "sitemap.json"


def save(ws, data):
    path = file_of(ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    temp.replace(path)
    return path


def load(ws):
    path = file_of(ws)
    if not path.exists():
        raise SitemapError("No sitemap for this workspace yet: run `oqa sitemap build` (it reads the menus and views, it clicks nothing).")
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schema_version", 0) > SCHEMA:
        raise SitemapError("This sitemap was written by a newer odoo-test: update the skill.")
    return data


# --- building -----------------------------------------------------------------------
def _short(text, size=80):
    text = " ".join(str(text or "").split())
    return text if len(text) <= size else text[:size - 1] + "…"


def buttons_of_arch(arch):
    """The clickable controls of a view definition that act: [{label, name, type, [invisible], [stat]}]; None when the definition is not readable XML."""
    try:
        root = ET.fromstring(arch.strip())
    except (ET.ParseError, AttributeError):
        return None
    out = []
    for element in root.iter():
        kind = element.get("type")
        if element.tag not in ("button", "a") or kind not in ("action", "object") + OPEN_ROW:
            continue
        item = {"name": element.get("name") or "", "type": kind, "label": _short(element.get("string") or "".join(element.itertext()) or element.get("title") or element.get("name"), 60)}
        if element.get("invisible"):
            item["invisible"] = _short(element.get("invisible"))
        if "oe_stat_button" in (element.get("class") or ""):
            item["stat"] = True
        out.append(item)
    return out


def disabled_of_arch(arch):
    """What the view switches off in its root tag (`create="0"` hides New, `edit="0"` makes it read only): ['create', 'edit', ...]."""
    try:
        root = ET.fromstring(arch.strip())
    except (ET.ParseError, AttributeError):
        return []
    return [name for name in ("create", "edit", "delete", "duplicate") if (root.get(name) or "").lower() in ("0", "false")]


def fields_of_arch(arch, info):
    """The fields a view shows (not the columns inside a one2many, those belong to that field): [{name, label, type, required, readonly, [relation, options, invisible, widget]}].
    `info` is what Odoo says about the model's fields (label, type, required, selection values); the view adds its own `string`, `required`, `readonly`, `widget`."""
    try:
        root = ET.fromstring(arch.strip())
    except (ET.ParseError, AttributeError):
        return []
    out, stack = [], [root]
    while stack:
        element = stack.pop()
        if element.tag == "field" and element.get("name"):
            name, meta = element.get("name"), info.get(element.get("name")) or {}
            hidden = element.get("invisible")
            if hidden not in ("1", "True", "true"):
                item = {"name": name, "label": _short(element.get("string") or meta.get("string") or name, 60), "type": meta.get("type") or "",
                        "required": _flag(element.get("required"), meta.get("required")), "readonly": _flag(element.get("readonly"), meta.get("readonly"))}
                if meta.get("relation"):
                    item["relation"] = meta["relation"]
                if meta.get("selection"):
                    item["options"] = [f"{v}={l}" for v, l in meta["selection"][:12]]
                if hidden:
                    item["invisible"] = _short(hidden)
                if element.get("widget"):
                    item["widget"] = element.get("widget")
                out.append(item)
            continue                                             # what is inside a field (a sub-list, a sub-form) is not a field of this view
        stack.extend(reversed(list(element)))
    return out


def _flag(from_view, from_model):
    """True / False, or the expression the view uses (the field is only sometimes required or read only)."""
    if from_view in (None, ""):
        return bool(from_model)
    return True if from_view in ("1", "True", "true") else False if from_view in ("0", "False", "false") else _short(from_view)


def _action_ref(name, by_name):
    """(action model or None, id) a button's `name` points at: `%(module.xmlid)d`, `module.xmlid` or a number; None when it is a method name."""
    name = (name or "").strip()
    wrapped = re.fullmatch(r"%\((.+)\)d", name)
    name = wrapped.group(1) if wrapped else name
    if name.isdigit():
        return (None, int(name))
    return by_name.get(name)


def build(rpc, profile="", version=""):
    nodes, edges, unresolved = {}, [], []

    def edge(source, target, kind, **extra):
        edges.append({"from": source, "to": target, "kind": kind, **{k: v for k, v in extra.items() if v}})

    by_id, by_name = {}, {}
    try:
        for row in rpc("ir.model.data", "search_read", [["|", ["model", "=", "ir.ui.menu"], ["model", "=like", "ir.actions.%"]]], {"fields": ["module", "name", "model", "res_id"]}):
            by_id[(row["model"], row["res_id"])] = f"{row['module']}.{row['name']}"
            by_name[f"{row['module']}.{row['name']}"] = (row["model"], row["res_id"])
    except AssertionError as error:
        unresolved.append({"where": "external ids", "why": f"not readable with this profile, so menus and actions have no external id and `%(module.name)d` buttons cannot be followed: {error}"})

    menus = rpc("ir.ui.menu", "search_read", [[]], {"fields": ["name", "parent_id", "action", "sequence"], "order": "sequence,id"})
    menu_by_id = {m["id"]: m for m in menus}

    def trail(menu):
        names, seen = [], set()
        while menu and menu["id"] not in seen:
            seen.add(menu["id"])
            names.append(menu["name"])
            menu = menu_by_id.get(menu["parent_id"][0]) if menu["parent_id"] else None
        return " / ".join(reversed(names))

    pending = {}                                         # action model (None: only the number is known) -> ids still to read
    number_key = {}                                      # an action known by number -> its node key, once its type is read
    waiting = []                                         # (view key, button) whose action is known by number only

    def want(model, res_id):
        if f"{KINDS[model]}:{res_id}" not in nodes:
            pending.setdefault(model, set()).add(res_id)

    for menu in menus:
        key = f"menu:{menu['id']}"
        nodes[key] = {"kind": "menu", "label": trail(menu), "name": menu["name"], **({"xmlid": by_id[("ir.ui.menu", menu["id"])]} if ("ir.ui.menu", menu["id"]) in by_id else {})}
        if menu["parent_id"]:
            edge(f"menu:{menu['parent_id'][0]}", key, "menu")
    for menu in menus:
        model, _, res_id = str(menu["action"] or "").partition(",")
        if res_id.isdigit() and model in KINDS:
            want(model, int(res_id))
            edge(f"menu:{menu['id']}", f"{KINDS[model]}:{res_id}", "opens")

    def settle(ids_by_model):
        """Read the actions asked for and make their nodes (an action known by number only has its type read first). Returns [(key, row)]."""
        unknown = ids_by_model.pop(None, set())
        if unknown:
            for row in rpc("ir.actions.actions", "read", [sorted(unknown)], {"fields": ["type"]}):
                if row["type"] in KINDS:
                    ids_by_model.setdefault(row["type"], set()).add(row["id"])
                    number_key[row["id"]] = f"{KINDS[row['type']]}:{row['id']}"
        made = []
        for model, ids in ids_by_model.items():
            kind = KINDS[model]
            ids = sorted(i for i in ids if f"{kind}:{i}" not in nodes)
            for row in (rpc(model, "read", [ids], {"fields": ACTION_FIELDS[kind]}) if ids else []):
                key = f"{kind}:{row['id']}"
                node = {"kind": "action", "type": kind, "label": row.get("name") or key}
                if (model, row["id"]) in by_id:
                    node["xmlid"] = by_id[(model, row["id"])]
                if kind == "window":
                    node.update(model=row["res_model"], target=row.get("target") or "current")
                elif kind in ("server", "report"):
                    node["model"] = row.get("model_name") or row.get("model") or ""
                nodes[key] = node
                made.append((key, row))
        return made

    def read_views(key, row):
        model = row["res_model"]
        wanted = [[v[0], v[1]] for v in (row.get("views") or []) if v[1] != "search"] or [[False, t] for t in (row.get("view_mode") or "list,form").split(",") if t != "search"]
        try:
            answer = rpc(model, "get_views", [], {"views": wanted, "options": {}})
            got = answer["views"]
        except AssertionError as error:
            unresolved.append({"where": key, "why": f"the views of {model} could not be read: {_short(error, 160)}"})
            return
        nodes[key]["views"] = sorted(got)
        known = (answer.get("models") or {}).get(model) or {}
        known = known["fields"] if isinstance(known.get("fields"), dict) else known       # Odoo 17 answers {field: info}, later versions {fields: {...}}
        made = {}
        for kind, info in got.items():
            view_key = f"view:{model}:{kind}:{info.get('id') or 0}"
            made[kind] = view_key
            edge(key, view_key, "shows")
            if view_key in nodes:
                continue
            found = buttons_of_arch(info.get("arch") or "")
            if found is None:
                unresolved.append({"where": view_key, "why": "the view definition is not readable XML"})
                found = []
            nodes[view_key] = {"kind": "view", "type": kind, "model": model, "label": f"{model} {kind}", "buttons": found, "fields": fields_of_arch(info.get("arch") or "", known), "disabled": disabled_of_arch(info.get("arch") or "")}
            for button in found:
                if button["type"] == "object":
                    counts["object"] += 1
                elif button["type"] == "action":
                    ref = _action_ref(button["name"], by_name)
                    if ref is None or (ref[0] is not None and ref[0] not in KINDS):
                        unresolved.append({"where": view_key, "why": f"button '{button['label']}' points at action '{button['name']}', which is not a known action"})
                    elif ref[0] is None:
                        waiting.append((view_key, button, ref[1]))
                        pending.setdefault(None, set()).add(ref[1])
                    else:
                        want(*ref)
                        edge(view_key, f"{KINDS[ref[0]]}:{ref[1]}", "button", label=button["label"], name=button["name"])
        for kind, view_key in made.items():
            if kind != "form" and "form" in made:
                edge(view_key, made["form"], "row", label="open a row")

    counts = {"object": 0}
    while pending and len(nodes) < MAX_NODES:
        batch, pending = pending, {}
        for key, row in settle(batch):
            if key.startswith("window:"):
                read_views(key, row)
        for view_key, button, number in list(waiting):
            if number in number_key:
                edge(view_key, number_key[number], "button", label=button["label"], name=button["name"])
                waiting.remove((view_key, button, number))
    truncated = bool(pending)
    for view_key, button, number in waiting:
        if number not in number_key:
            unresolved.append({"where": view_key, "why": f"button '{button['label']}' points at action number {number}, which could not be read"})
    for e in [e for e in edges if e["to"] not in nodes]:
        edges.remove(e)
        unresolved.append({"where": e["from"], "why": f"leads to {e['to']}, which could not be read"})
    kinds = [n["kind"] for n in nodes.values()]
    return {"schema_version": SCHEMA,
            "meta": {"generated": datetime.now(timezone.utc).isoformat(timespec="seconds"), "profile": profile, "odoo_version": version,
                     "menus": kinds.count("menu"), "actions": kinds.count("action"), "views": kinds.count("view"), "object_buttons": counts["object"], "truncated": truncated},
            "nodes": nodes, "edges": edges, "unresolved": unresolved}


# --- reading ------------------------------------------------------------------------
def age_days(data):
    try:
        return (datetime.now(timezone.utc) - datetime.fromisoformat(data["meta"]["generated"])).days
    except (KeyError, ValueError):
        return 0


def header(data):
    m = data["meta"]
    old = age_days(data)
    return (f"sitemap of {m.get('profile') or '?'} on Odoo {m.get('odoo_version') or '?'}, built {m['generated'][:10]}: {m['menus']} menus, {m['actions']} actions, {m['views']} views"
            + (f". It is {old} days old: if a module was installed since, run `oqa sitemap build` again." if old >= STALE_DAYS else "")
            + (". Not complete: the crawl stopped at its size limit." if m.get("truncated") else ""))


def _words(text):
    return re.findall(r"\w+", str(text).lower())


def find(data, text, limit=12, kind=""):
    """[(key, node, reason)] for the screens whose label, model, external id or button names contain the words of `text`, best first."""
    words = _words(text)
    if not words:
        return []
    order = {"menu": 0, "action": 1, "view": 2}
    scored = []
    for key, node in data["nodes"].items():
        if kind and node["kind"] != kind:
            continue
        hay = " ".join(str(node.get(k, "")) for k in ("label", "model", "xmlid")).lower()
        reason, score = "", sum(1 for w in words if w in hay)
        for button in node.get("buttons") or []:
            if all(w in f"{button['label']} {button['name']}".lower() for w in words):      # a button counts only when it holds every word asked
                reason, score = f"button '{button['label']}'" + (f" [name={button['name']}]" if button["name"] else ""), max(score, len(words))
                break
        if score:
            scored.append((-score, order[node["kind"]], key, node, reason))
    scored.sort(key=lambda s: s[:3])
    return [(key, node, reason) for _, _, key, node, reason in scored[:limit]]


def resolve(data, text):
    """The node key `text` stands for: a key itself, or the best match of `find`."""
    if text in data["nodes"]:
        return text
    found = find(data, text, 1)
    return found[0][0] if found else None


def _out(data):
    table = {}
    for e in data["edges"]:
        table.setdefault(e["from"], []).append(e)
    return table


def views_of(data, key):
    """The view nodes reached from a menu, an action or a view itself."""
    out = _out(data)
    node = data["nodes"][key]
    if node["kind"] == "view":
        return [key]
    if node["kind"] == "menu":
        return [v for e in out.get(key, []) if e["kind"] == "opens" for v in views_of(data, e["to"])]
    return [e["to"] for e in out.get(key, []) if e["kind"] == "shows"]


def path(data, source, target):
    """The shortest way from the screen `source` (any menu root when empty) to `target`: [edge], or None."""
    goal = resolve(data, target)
    if goal is None:
        raise SitemapError(f"nothing in the sitemap matches '{target}'")
    start = resolve(data, source) if source else None
    if source and start is None:
        raise SitemapError(f"nothing in the sitemap matches '{source}'")
    out = _out(data)
    children = {e["to"] for e in data["edges"] if e["kind"] == "menu"}
    roots = [start] if start else [k for k, n in data["nodes"].items() if n["kind"] == "menu" and k not in children]
    seen, queue = set(roots), deque((r, []) for r in roots)
    while queue:
        key, way = queue.popleft()
        if key == goal:
            return way
        for e in out.get(key, []):
            if e["to"] not in seen:
                seen.add(e["to"])
                queue.append((e["to"], way + [e]))
    return None


# --- text for people and for the assistant (CLI and MCP share these) -----------------------------
def selector_of(kind, name, view_type=""):
    """A selector to start from (a hint: only `oqa try` proves it on the page)."""
    if kind == "button":
        return f"button[name='{name}']"
    if kind == "menu":
        return f"[data-menu-xmlid='{name}']"
    return f"th[data-name='{name}']" if view_type == "list" else f".o_field_widget[name='{name}']"


def describe(key, node):
    if node["kind"] == "menu":
        return f"menu {node['label']}" + (f"  [{node['xmlid']}]" if node.get("xmlid") else "")
    if node["kind"] == "action":
        return f"{node['type']} action \"{node['label']}\"" + (f" on {node['model']}" if node.get("model") else "") + (" (a dialog)" if node.get("target") == "new" else "") + (f"  [{node['xmlid']}]" if node.get("xmlid") else "")
    return f"{node['type']} view of {node['model']}"


def _cut(lines, limit, hint):
    """At most `limit` lines, and when some are left out, say how to narrow the question instead of printing them."""
    return lines if len(lines) <= limit else lines[:limit] + [f"    ... {len(lines) - limit} more not shown: narrow with {hint}"]


def _views(data, key, view=""):
    """The views of a screen, only the ones of type `view` when it is given (`form`, `list`, ...)."""
    return [v for v in views_of(data, key) if not view or data["nodes"][v]["type"] == view]


def _head(data, key, node, view_key, extra=""):
    view = data["nodes"][view_key]
    return f"  {describe(view_key, view)}  ({view_key})" + (f"  NOT possible here: {', '.join(view['disabled'])}" if view.get("disabled") else "") + extra


def text_find(data, text, limit=12, kind=""):
    found = find(data, text, limit, kind)
    lines = [f"{key}  {describe(key, node)}" + (f"  <- {reason}" if reason else "") for key, node, reason in found]
    return "\n".join([header(data), *(lines or [f"nothing matches '{text}'" + (f" among {kind}s" if kind else "") + ". Try fewer or other words (a menu name, a model, a button label)."])])


def text_buttons(data, screen, view="", match="", limit=30):
    """The buttons of a screen's views, where each leads. `view`: only that view type; `match`: only buttons whose label or name contains it."""
    key = resolve(data, screen)
    if key is None:
        return f"nothing in the sitemap matches '{screen}'"
    out, node = _out(data), data["nodes"][key]
    lines = [f"{key}  {describe(key, node)}" + (f"  {selector_of('menu', node['xmlid'])}" if node.get("xmlid") and key.startswith("menu:") else "")]
    for view_key in _views(data, key, view):
        vnode = data["nodes"][view_key]
        lines.append(_head(data, key, node, view_key))
        goes = {e.get("name"): e["to"] for e in out.get(view_key, []) if e["kind"] == "button"}
        rows = []
        for b in vnode["buttons"]:
            if b["type"] in OPEN_ROW or (match and match.lower() not in f"{b['label']} {b['name']}".lower()):
                continue
            where = (f"-> {describe(goes[b['name']], data['nodes'][goes[b['name']]])}" if b["name"] in goes else
                     "-> unknown until it is run (it calls a method)" if b["type"] == "object" else "-> not resolved")
            rows.append(f"    {'stat ' if b.get('stat') else ''}button \"{b['label']}\" [name={b['name']}] {selector_of('button', b['name'])}" + (f" shown only if: not ({b['invisible']})" if b.get("invisible") else "") + f" {where}")
        lines += _cut(rows, limit, "--match <word> or --view <type>")
        if not match:
            lines += [f"    clicking a row -> {describe(e['to'], data['nodes'][e['to']])}" for e in out.get(view_key, []) if e["kind"] == "row"]
    return "\n".join(lines)


def text_fields(data, screen, view="", match="", required=False, limit=40):
    """The fields of a screen, with a selector to start from, required / read only and the values of a selection. By default only the form (the list repeats it);
    `view`: another view type; `match`: only fields whose name or label contains it; `required`: only fields that are (or can be) required. A field shown twice is listed once."""
    key = resolve(data, screen)
    if key is None:
        return f"nothing in the sitemap matches '{screen}'"
    node = data["nodes"][key]
    views = _views(data, key, view)
    if not view and any(data["nodes"][v]["type"] == "form" for v in views):
        views = [v for v in views if data["nodes"][v]["type"] == "form"]
    lines = [f"{key}  {describe(key, node)}"]
    for view_key in views:
        vnode = data["nodes"][view_key]
        lines.append(_head(data, key, node, view_key))
        rows, seen = [], set()
        for f in vnode.get("fields") or []:
            if f["name"] in seen or (match and match.lower() not in f"{f['name']} {f['label']}".lower()) or (required and not f["required"]):
                continue
            seen.add(f["name"])
            flags = [w for w, on in (("required", f["required"]), ("readonly", f["readonly"])) if on is True] + [f"{w} if: {on}" for w, on in (("required", f["required"]), ("readonly", f["readonly"])) if isinstance(on, str)]
            rows.append(f"    {f['name']} \"{f['label']}\" {f['type']}" + (f" -> {f['relation']}" if f.get("relation") else "") + (f" [{', '.join(flags)}]" if flags else "")
                        + (f" shown only if: not ({f['invisible']})" if f.get("invisible") else "") + (f" widget={f['widget']}" if f.get("widget") else "") + f"  {selector_of('field', f['name'], vnode['type'])}"
                        + (f"  options: {', '.join(f['options'])}" if f.get("options") else ""))
        lines += _cut(rows, limit, "--match <word>, --required or --view <type>")
    return "\n".join(lines)


def text_path(data, source, target):
    way = path(data, source, target)
    goal = resolve(data, target)
    if way is None:
        return f"no way found to {describe(goal, data['nodes'][goal])} from {source or 'any menu'}: it may be reached only by a button that calls a method."
    first = way[0]["from"] if way else goal
    lines = [f"1. start at {describe(first, data['nodes'][first])}"]
    for number, e in enumerate(way, 2):
        node = data["nodes"][e["to"]]
        how = {"menu": "open the submenu" + (f" {selector_of('menu', node['xmlid'])}" if node.get("xmlid") else ""), "opens": "it opens", "shows": "it shows", "row": "click a row",
               "button": f"click the button \"{e.get('label', '')}\" {selector_of('button', e.get('name', ''))}"}[e["kind"]]
        lines.append(f"{number}. {how}: {describe(e['to'], data['nodes'][e['to']])}")
    return "\n".join(lines)
