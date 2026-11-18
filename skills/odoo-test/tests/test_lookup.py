"""Read-only lookups through the ORM (o.models, o.fields, o.xmlid, o.find, o.records): only reading methods go out, and the answers are shaped for a script."""
import pytest

from oqa.odoo_ui import Odoo, MAX_ROWS


class Reply:
    def __init__(self, body):
        self.body = body

    def json(self):
        return self.body


class FakePage:
    """Records what is posted and answers from `answers[(model, method)]`."""

    def __init__(self, answers):
        self.answers, self.sent = answers, []
        self.request = self

    def set_default_timeout(self, ms):
        pass

    def post(self, url, data):
        params = data["params"]
        self.sent.append((url, params))
        answer = self.answers[(params["model"], params["method"])]
        return Reply({"error": {"message": answer["error"]}} if isinstance(answer, dict) and "error" in answer else {"result": answer})


def odoo(answers):
    page = FakePage(answers)
    return Odoo(page, "http://odoo.test/"), page


def test_only_reading_methods_are_sent():
    o, page = odoo({})
    for method in ("write", "create", "unlink", "action_confirm", "execute_kw"):
        with pytest.raises(PermissionError, match="only reads"):
            o.rpc("sale.order", method, [[1]])
    assert page.sent == []


def test_models_are_listed_by_technical_name_and_label():
    o, page = odoo({("ir.model", "search_read"): [{"model": "sale.order", "name": "Sales Order"}, {"model": "sale.order.line", "name": "Sales Order Line"}]})
    assert o.models("sale.order") == [("sale.order", "Sales Order"), ("sale.order.line", "Sales Order Line")]
    url, params = page.sent[0]
    assert url == "http://odoo.test/web/dataset/call_kw/ir.model/search_read"
    assert params["args"] == [["|", ["model", "ilike", "sale.order"], ["name", "ilike", "sale.order"]]]


def test_model_exists_counts_the_exact_name():
    o, page = odoo({("ir.model", "search_count"): 1})
    assert o.model_exists("sale.order") is True
    assert page.sent[0][1]["args"] == [[["model", "=", "sale.order"]]]
    o, _ = odoo({("ir.model", "search_count"): 0})
    assert o.model_exists("sale.orders") is False


def test_fields_match_on_name_or_label_and_drop_empty_attributes():
    info = {"partner_id": {"string": "Customer", "type": "many2one", "relation": "res.partner", "required": True, "readonly": False},
            "name": {"string": "Order Reference", "type": "char", "relation": None, "required": True, "readonly": False},
            "note": {"string": "Terms", "type": "html", "relation": None, "required": False, "readonly": False}}
    o, _ = odoo({("sale.order", "fields_get"): info})
    assert [f["name"] for f in o.fields("sale.order", "customer")] == ["partner_id"]
    assert o.fields("sale.order", "customer")[0]["relation"] == "res.partner"
    assert "readonly" not in o.fields("sale.order", "terms")[0]
    assert [f["name"] for f in o.fields("sale.order", "ref")] == ["name"]


def test_xmlid_gives_model_and_id_or_none():
    o, page = odoo({("ir.model.data", "search_read"): [{"model": "res.company", "res_id": 1}]})
    assert o.xmlid("base.main_company") == {"ref": "base.main_company", "model": "res.company", "res_id": 1}
    assert page.sent[0][1]["args"] == [[["module", "=", "base"], ["name", "=", "main_company"]]]
    o, _ = odoo({("ir.model.data", "search_read"): []})
    assert o.xmlid("base.nothing") is None


def test_find_returns_id_and_name_pairs():
    o, page = odoo({("res.partner", "name_search"): [[7, "Azure Interior"]]})
    assert o.find("res.partner", "Azure") == [(7, "Azure Interior")]
    assert page.sent[0][1]["kwargs"] == {"name": "Azure", "limit": 5}


def test_records_are_capped_and_errors_say_what_the_server_said():
    o, page = odoo({("res.partner", "search_read"): []})
    o.records("res.partner", [], ["name"], limit=100000)
    assert page.sent[0][1]["kwargs"]["limit"] == MAX_ROWS
    o, _ = odoo({("res.partner", "search_read"): {"error": "Invalid field 'nme'"}})
    with pytest.raises(AssertionError, match="res.partner.search_read: Invalid field 'nme'"):
        o.records("res.partner", [], ["nme"])
