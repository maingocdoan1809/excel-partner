"""A small MCP server (stdio, JSON-RPC lines) that lets the assistant ask the workspace's sitemap before `oqa try`: where is a screen, which buttons it has,
how to get from one screen to another. It only reads `knowledge/sitemap.json` (built by `oqa sitemap build`); it opens no browser and needs no package."""
import json
import sys

from . import paths, sitemap, workspace

PROTOCOL = "2024-11-05"
TOOLS = [
    {"name": "sitemap_find", "description": "Find screens of the Odoo under test (menus, actions, views) by words from a menu name, a model, an external id or a button label. Use it before oqa explore / oqa try.",
     "inputSchema": {"type": "object", "properties": {"text": {"type": "string"}, "kind": {"type": "string", "enum": ["menu", "action", "view"], "description": "only this kind of screen"}, "limit": {"type": "integer", "default": 12}}, "required": ["text"]}},
    {"name": "sitemap_buttons", "description": "The buttons of a screen (a key from sitemap_find, or words), with their technical [name=...], a selector to start from, when they show, and where they lead.",
     "inputSchema": {"type": "object", "properties": {"screen": {"type": "string"}, "view": {"type": "string", "description": "only this view type: form, list, kanban..."}, "match": {"type": "string", "description": "only buttons whose label or name contains this"}, "limit": {"type": "integer", "default": 30}}, "required": ["screen"]}},
    {"name": "sitemap_fields", "description": "The fields of a screen's views: technical name, label, type, required / read only, the values of a selection, and a selector to start from. Use it to know what a form needs before filling it. Start with required=true, then match=<word> for the fields you need.",
     "inputSchema": {"type": "object", "properties": {"screen": {"type": "string"}, "required": {"type": "boolean", "description": "only fields that are or can be required"}, "match": {"type": "string", "description": "only fields whose name or label contains this"}, "view": {"type": "string", "description": "default: the form"}, "limit": {"type": "integer", "default": 40}}, "required": ["screen"]}},
    {"name": "sitemap_path", "description": "The shortest way from one screen to another: menu, action, view, button. Leave `source` empty to start from the menus.",
     "inputSchema": {"type": "object", "properties": {"source": {"type": "string"}, "target": {"type": "string"}}, "required": ["target"]}},
]
for tool in TOOLS:
    tool["inputSchema"]["properties"]["lang"] = {"type": "string", "enum": ["en", "ja", "en_US", "ja_JP"], "description": "Response language; defaults to the workspace language. Odoo labels remain searchable in both English and Japanese."}


def call(ws, name, arguments):
    """The text a tool answers (shared with `oqa sitemap find|buttons|path`); raises SitemapError / PathError when it cannot."""
    ws = paths.resolve_workspace(ws)
    lang = sitemap.output_language(arguments.get("lang") or workspace.load(ws).get("language"))
    data = sitemap.load(ws, lang)
    if name == "sitemap_find":
        return sitemap.text_find(data, arguments["text"], int(arguments.get("limit") or 12), arguments.get("kind") or "", lang=lang)
    if name == "sitemap_buttons":
        return sitemap.text_buttons(data, arguments["screen"], arguments.get("view") or "", arguments.get("match") or "", int(arguments.get("limit") or 30), lang=lang)
    if name == "sitemap_fields":
        return sitemap.text_fields(data, arguments["screen"], arguments.get("view") or "", arguments.get("match") or "", bool(arguments.get("required")), int(arguments.get("limit") or 40), lang=lang)
    if name == "sitemap_path":
        return sitemap.text_path(data, arguments.get("source") or "", arguments["target"], lang=lang)
    raise sitemap.SitemapError(f"no such tool: {name}")


def handle(message, ws=None):
    """The answer to one JSON-RPC message, or None for a notification."""
    method, ident, params = message.get("method"), message.get("id"), message.get("params") or {}
    if ident is None:
        return None
    reply = {"jsonrpc": "2.0", "id": ident}
    if method == "initialize":
        reply["result"] = {"protocolVersion": params.get("protocolVersion") or PROTOCOL, "capabilities": {"tools": {}}, "serverInfo": {"name": "oqa-sitemap", "version": "1"}}
    elif method == "tools/list":
        reply["result"] = {"tools": TOOLS}
    elif method == "tools/call":
        try:
            text, bad = call(ws, params.get("name"), params.get("arguments") or {}), False
        except (sitemap.SitemapError, paths.PathError, KeyError) as error:
            text, bad = f"{type(error).__name__}: {error}" if isinstance(error, KeyError) else str(error), True
        reply["result"] = {"content": [{"type": "text", "text": text}], "isError": bad}
    elif method == "ping":
        reply["result"] = {}
    else:
        reply["error"] = {"code": -32601, "message": f"unknown method {method}"}
    return reply


def serve(ws=None, stdin=None, stdout=None):
    stdin, stdout = stdin or sys.stdin, stdout or sys.stdout
    for line in stdin:
        if not line.strip():
            continue
        try:
            reply = handle(json.loads(line), ws)
        except json.JSONDecodeError:
            reply = {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "not JSON"}}
        if reply is not None:
            stdout.write(json.dumps(reply, ensure_ascii=False) + "\n")
            stdout.flush()
