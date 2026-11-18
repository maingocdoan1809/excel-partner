"""The sitemap: built from read-only answers of Odoo (menus, actions, view definitions), saved per workspace, searched by the CLI and the MCP server."""
import io
import json

import pytest

from oqa import cli, mcp_server, paths, sitemap
from oqa.odoo_ui import READ_METHODS

SALE_LIST = '<list create="0" delete="false"><field name="name"/><field name="state"/><button name="action_quotation_send" type="object" string="Send"/></list>'
SALE_FORM = ('<form><header><button name="action_confirm" type="object" string="Confirm" invisible="state != \'draft\'"/>'
             '<button name="%(sale.action_cancel_wizard)d" type="action" string="Cancel"/></header>'
             '<field name="partner_id" required="state == \'sale\'"/><field name="note" invisible="state == \'draft\'"/><field name="secret" invisible="1"/>'
             '<field name="line_ids"><list><field name="inner"/></list></field>'
             '<div class="oe_button_box"><button name="action_view_invoice" type="object" class="oe_stat_button"><field name="invoice_count"/> Invoices</button></div></form>')
SALE_FIELDS = {"name": {"string": "Reference", "type": "char", "required": True}, "state": {"string": "Status", "type": "selection", "selection": [["draft", "Quotation"], ["sale", "Sales Order"]]},
               "partner_id": {"string": "Customer", "type": "many2one", "relation": "res.partner", "required": True}, "note": {"string": "Note", "type": "text"}}
WIZARD_FORM = '<form><footer><button name="action_cancel" type="object" string="Cancel order"/><button special="cancel" string="Discard"/></footer></form>'


class FakeRpc:
    """What a small Odoo answers to the calls `sitemap.build` makes."""

    def __init__(self, broken=()):
        self.calls, self.broken = [], broken

    def __call__(self, model, method, args=None, kwargs=None):
        self.calls.append((model, method))
        if (model, method) == ("ir.model.data", "search_read"):
            return [{"module": "sale", "name": "menu_orders", "model": "ir.ui.menu", "res_id": 2},
                    {"module": "sale", "name": "action_orders", "model": "ir.actions.act_window", "res_id": 10},
                    {"module": "sale", "name": "action_cancel_wizard", "model": "ir.actions.act_window", "res_id": 11}]
        if (model, method) == ("ir.ui.menu", "search_read"):
            return [{"id": 1, "name": "Sales", "parent_id": False, "action": False}, {"id": 2, "name": "Orders", "parent_id": [1, "Sales"], "action": "ir.actions.act_window,10"},
                    {"id": 3, "name": "Reports", "parent_id": [1, "Sales"], "action": "ir.actions.act_window,12"}]
        if (model, method) == ("ir.actions.act_window", "read"):
            rows = {10: {"id": 10, "name": "Quotations", "res_model": "sale.order", "views": [[False, "list"], [False, "form"]], "target": "current"},
                    11: {"id": 11, "name": "Cancel order", "res_model": "sale.order.cancel", "views": [[False, "form"]], "target": "new"},
                    12: {"id": 12, "name": "Hidden", "res_model": "sale.secret", "views": [[False, "list"]], "target": "current"}}
            return [rows[i] for i in args[0]]
        if method == "get_views":
            if model == "sale.secret":
                raise AssertionError("AccessError: not allowed")
            arch = {"sale.order": {"list": SALE_LIST, "form": SALE_FORM}, "sale.order.cancel": {"form": WIZARD_FORM}}[model]
            return {"views": {kind: {"id": index + 1, "arch": text} for index, (kind, text) in enumerate(arch.items())}, "models": {model: {"fields": SALE_FIELDS} if model == "sale.order" else {}}}
        raise AssertionError(f"unexpected call {model}.{method}")


@pytest.fixture
def data():
    return sitemap.build(FakeRpc(), profile="Admin", version="18.0")


def test_menus_actions_and_views_are_linked(data):
    nodes, edges = data["nodes"], data["edges"]
    assert nodes["menu:2"]["label"] == "Sales / Orders" and nodes["menu:2"]["xmlid"] == "sale.menu_orders"
    assert nodes["window:11"]["target"] == "new"
    kinds = {(e["from"], e["to"], e["kind"]) for e in edges}
    assert ("menu:1", "menu:2", "menu") in kinds and ("menu:2", "window:10", "opens") in kinds and ("window:10", "view:sale.order:form:2", "shows") in kinds
    assert ("view:sale.order:list:1", "view:sale.order:form:2", "row") in kinds


def test_a_button_that_opens_an_action_is_followed_and_a_method_button_is_not(data):
    edge = next(e for e in data["edges"] if e["kind"] == "button")
    assert (edge["from"], edge["to"], edge["label"]) == ("view:sale.order:form:2", "window:11", "Cancel")
    assert data["meta"]["object_buttons"] == 4                 # send, confirm, view invoices, cancel order
    assert "view:sale.order.cancel:form:1" in data["nodes"]    # the dialog's own buttons are read too


def test_what_could_not_be_read_is_reported_not_hidden(data):
    assert any("sale.secret" in u["why"] and u["where"] == "window:12" for u in data["unresolved"])
    assert data["meta"]["views"] == 3 and data["meta"]["truncated"] is False


def test_the_crawl_only_reads():
    rpc = FakeRpc()
    sitemap.build(rpc)
    assert {method for _, method in rpc.calls} <= {"search_read", "read", "get_views"}
    assert "get_views" in READ_METHODS


def test_find_matches_menu_model_and_button_words(data):
    top = lambda text: sitemap.find(data, text)[0]
    assert top("sales orders")[0] == "menu:2"
    assert top("confirm")[0] == "view:sale.order:form:2" and "action_confirm" in top("confirm")[2]
    assert sitemap.find(data, "xyzzy") == []


def test_buttons_of_a_menu_say_where_each_leads(data):
    text = sitemap.text_buttons(data, "Orders")
    assert 'button "Cancel" [name=%(sale.action_cancel_wizard)d] button[name=' in text and "(a dialog)" in text
    assert 'button "Confirm" [name=action_confirm] button[name=\'action_confirm\'] shown only if: not (state != \'draft\')' in text and "calls a method" in text
    assert 'stat button' in text and "clicking a row" in text


def test_path_goes_from_the_menus_to_a_dialog(data):
    text = sitemap.text_path(data, "", "Cancel order")
    assert text.splitlines()[0].startswith("1. start at menu Sales")
    assert 'click the button "Cancel"' in text and text.endswith("(a dialog)  [sale.action_cancel_wizard]")
    assert [e["kind"] for e in sitemap.path(data, "", "window:11")] == ["menu", "opens", "shows", "button"]
    assert sitemap.path(data, "window:11", "menu:2") is None
    with pytest.raises(sitemap.SitemapError):
        sitemap.path(data, "", "nonexistent screen xyz")


def test_a_button_by_number_is_resolved_through_its_type():
    rpc = FakeRpc()
    original = rpc.__call__
    form = SALE_FORM.replace("%(sale.action_cancel_wizard)d", "11")

    def call(model, method, args=None, kwargs=None):
        if (model, method) == ("ir.actions.actions", "read"):
            return [{"id": i, "type": "ir.actions.act_window"} for i in args[0]]
        if method == "get_views" and model == "sale.order":
            return {"views": {"form": {"id": 2, "arch": form}}}
        return original(model, method, args, kwargs)
    result = sitemap.build(call)
    assert any(e["kind"] == "button" and e["to"] == "window:11" for e in result["edges"])


def test_saved_per_workspace_and_stale_warning(data, ws):
    path = sitemap.save("demo", data)
    assert path == paths.workspace_dir("demo") / "knowledge" / "sitemap.json" and sitemap.load("demo")["meta"]["profile"] == "Admin"
    assert "days old" not in sitemap.header(data)
    data["meta"]["generated"] = "2020-01-01T00:00:00+00:00"
    assert "days old" in sitemap.header(data)
    with pytest.raises(sitemap.SitemapError, match="oqa sitemap build"):
        sitemap.load("nope")


def test_the_mcp_server_answers_the_three_tools(data, ws):
    sitemap.save("demo", data)
    send = lambda **m: mcp_server.handle({"jsonrpc": "2.0", **m}, "demo")
    assert send(method="notifications/initialized") is None
    assert send(id=1, method="initialize", params={"protocolVersion": "2025-03-26"})["result"]["protocolVersion"] == "2025-03-26"
    assert [t["name"] for t in send(id=2, method="tools/list")["result"]["tools"]] == ["sitemap_find", "sitemap_buttons", "sitemap_fields", "sitemap_path"]
    found = send(id=3, method="tools/call", params={"name": "sitemap_find", "arguments": {"text": "confirm"}})["result"]
    assert found["isError"] is False and "action_confirm" in found["content"][0]["text"]
    assert "(a dialog)" in send(id=4, method="tools/call", params={"name": "sitemap_path", "arguments": {"target": "Cancel order"}})["result"]["content"][0]["text"]
    assert send(id=5, method="tools/call", params={"name": "sitemap_find", "arguments": {}})["result"]["isError"] is True
    assert send(id=6, method="nope")["error"]["code"] == -32601


def test_the_mcp_server_says_so_when_there_is_no_sitemap(ws):
    reply = mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "sitemap_find", "arguments": {"text": "x"}}}, "demo")
    assert reply["result"]["isError"] is True and "oqa sitemap build" in reply["result"]["content"][0]["text"]


def test_serve_reads_lines_and_writes_lines(data, ws):
    sitemap.save("demo", data)
    out = io.StringIO()
    mcp_server.serve("demo", io.StringIO(json.dumps({"jsonrpc": "2.0", "id": 1, "method": "ping"}) + "\n\nnot json\n"), out)
    first, second = (json.loads(x) for x in out.getvalue().splitlines())
    assert first["result"] == {} and second["error"]["code"] == -32700


def test_cli_find_and_config(data, ws, capsys):
    sitemap.save("demo", data)
    assert cli.main(["sitemap", "find", "confirm", "--ws", "demo"]) == 0
    assert "action_confirm" in capsys.readouterr().out
    assert cli.main(["sitemap", "path", "Cancel order", "--ws", "demo"]) == 0
    assert "click the button" in capsys.readouterr().out
    assert cli.main(["mcp", "--config", "--ws", "demo"]) == 0
    assert json.loads(capsys.readouterr().out)["mcpServers"]["oqa-sitemap"]["args"][-2:] == ["--ws", "demo"]
    assert cli.main(["sitemap", "find", "--ws", "demo"]) == 2


def test_fields_carry_type_requirement_options_and_a_selector(data):
    text = sitemap.text_fields(data, "window:10", view="list")
    assert 'name "Reference" char [required]' in text and "th[data-name='name']" in text                       # a list column
    assert "options: draft=Quotation, sale=Sales Order" in text
    text = sitemap.text_fields(data, "window:10", view="form")
    assert "options: draft=Quotation, sale=Sales Order" not in text
    assert 'partner_id "Customer" many2one -> res.partner [required if: state == \'sale\']  .o_field_widget[name=\'partner_id\']' in text.replace("  ", "  ")
    assert 'note "Note" text shown only if: not (state == \'draft\')' in text
    assert "secret" not in text and "inner" not in text                                                            # a hidden field and a sub-list column are not fields of the form


def test_fields_also_read_the_older_answer_shape():
    arch = '<form><field name="a"/></form>'
    assert sitemap.fields_of_arch(arch, {"a": {"string": "A", "type": "char"}})[0]["label"] == "A"


def test_buttons_and_path_show_selectors(data):
    assert "button[name='action_confirm']" in sitemap.text_buttons(data, "Orders")
    assert "[data-menu-xmlid='sale.menu_orders']" in sitemap.text_buttons(data, "menu:2") and "button[name='%(sale.action_cancel_wizard)d']" in sitemap.text_path(data, "", "Cancel order")


def test_a_view_that_switches_off_new_or_edit_says_so(data):
    assert data["nodes"]["view:sale.order:list:1"]["disabled"] == ["create", "delete"] and data["nodes"]["view:sale.order:form:2"]["disabled"] == []
    assert "NOT possible here: create, delete" in sitemap.text_buttons(data, "Orders") and "NOT possible here: create, delete" in sitemap.text_fields(data, "Orders", view="list")


def test_narrowing_a_question_prints_less(data):
    everything = sitemap.text_fields(data, "window:10", view="list")
    assert "list view" in everything and "form view" not in everything
    assert "list view" not in sitemap.text_fields(data, "window:10")                                              # the form by default: the list repeats it
    only_required = sitemap.text_fields(data, "window:10", required=True)
    assert "partner_id" in only_required and "note" not in only_required
    assert "partner_id" in sitemap.text_fields(data, "window:10", match="cust") and "note" not in sitemap.text_fields(data, "window:10", match="cust")
    cut = sitemap.text_fields(data, "window:10", limit=1)
    assert "more not shown: narrow with --match" in cut
    assert sitemap.text_buttons(data, "window:10", match="confirm").count("button ") == 1
    assert "list view" not in sitemap.text_buttons(data, "window:10", view="form")
    assert all(line.startswith("window:") or "action" in line for line in sitemap.text_find(data, "orders", kind="action").splitlines()[1:])
    assert "among menus" in sitemap.text_find(data, "xyzzy", kind="menu")


def test_a_field_shown_twice_is_listed_once(data):
    data["nodes"]["view:sale.order:form:2"]["fields"] += [dict(f) for f in data["nodes"]["view:sale.order:form:2"]["fields"]]
    text = sitemap.text_fields(data, "window:10")
    assert text.count("    partner_id ") == 1


def test_the_cli_takes_the_narrowing_options(data, ws, capsys):
    sitemap.save("demo", data)
    assert cli.main(["sitemap", "fields", "Orders", "--required", "--ws", "demo"]) == 0
    out = capsys.readouterr().out
    assert "partner_id" in out and "note" not in out
    assert cli.main(["sitemap", "buttons", "Orders", "--match", "cancel", "--view", "form", "--ws", "demo"]) == 0
    assert capsys.readouterr().out.count("button ") == 1
