"""Read Odoo translations, search across languages, and render English/Japanese CLI/MCP answers."""
from contextlib import contextmanager
from copy import deepcopy

import pytest

from oqa import cli, mcp_server, sitemap, workspace
from test_sitemap import FakeRpc


class BilingualRpc(FakeRpc):
    def __init__(self, active=("en_US", "ja_JP")):
        super().__init__()
        self.active, self.contexts = active, []

    def __call__(self, model, method, args=None, kwargs=None):
        self.contexts.append(deepcopy((kwargs or {}).get("context", {})))
        if (model, method) == ("res.lang", "search_read"):
            self.calls.append((model, method))
            return [{"code": code} for code in self.active]
        result = deepcopy(super().__call__(model, method, args, kwargs))
        if (kwargs or {}).get("context", {}).get("lang") != "ja_JP":
            return result
        labels = {"Sales": "販売", "Orders": "注文", "Reports": "レポート", "Quotations": "見積", "Cancel order": "注文取消"}
        if model in ("ir.ui.menu", "ir.actions.act_window"):
            for row in result:
                row["name"] = labels.get(row["name"], row["name"])
        if method == "get_views":
            for view in result["views"].values():
                for english, japanese in {"Confirm": "確認", "Cancel": "取消", "Send": "送信", "Invoices": "請求書"}.items():
                    view["arch"] = view["arch"].replace(english, japanese)
            fields = result["models"][model].get("fields", {})
            for field in fields.values():
                field["string"] = {"Customer": "顧客", "Reference": "参照", "Status": "状態", "Note": "備考"}.get(field["string"], field["string"])
                if field.get("selection"):
                    field["selection"] = [[value, {"Quotation": "見積", "Sales Order": "販売注文"}.get(label, label)] for value, label in field["selection"]]
        return result


@pytest.fixture
def bilingual():
    return sitemap.build(BilingualRpc(), profile="Admin", version="19.0")


def test_build_collects_translations_without_changing_permissions_or_structure(bilingual):
    assert bilingual["meta"]["languages"] == ["en_US", "ja_JP"]
    assert bilingual["nodes"]["menu:2"]["labels"] == {"en_US": "Sales / Orders", "ja_JP": "販売 / 注文"}
    assert bilingual["nodes"]["window:10"]["labels"]["ja_JP"] == "見積"
    assert bilingual["meta"]["object_buttons"] == 4 and bilingual["meta"]["views"] == 3
    assert next(e for e in bilingual["edges"] if e["kind"] == "button")["labels"]["ja_JP"] == "取消"
    assert any(item.get("language") == "ja_JP" for item in bilingual["unresolved"])


def test_stat_button_labels_use_child_fields_and_ignore_formatting_whitespace():
    arch = '''<form><button name="action_show_groups" type="object" class="oe_stat_button">
        <field name="groups_count" string="グループ" widget="statinfo"/>
    </button><button name="action_show_accesses" type="object" class="oe_stat_button">
        <field name="accesses_count" widget="statinfo"/>
    </button><button name="icon_action" type="object" title="Details">   </button></form>'''
    buttons = sitemap.buttons_of_arch(arch, {"accesses_count": {"string": "アクセス権"}})
    assert [b["label"] for b in buttons] == ["グループ", "アクセス権", "Details"]
    assert buttons[0]["name"] == "action_show_groups" and buttons[0]["stat"] is True


def test_language_context_is_explicit_and_does_not_mutate_callers():
    rpc = BilingualRpc()
    context = {"allowed_company_ids": [3], "tz": "Asia/Tokyo", "lang": "vi_VN"}
    sitemap.build(rpc, context=context)
    assert context == {"allowed_company_ids": [3], "tz": "Asia/Tokyo", "lang": "vi_VN"}
    assert all(c["allowed_company_ids"] == [3] and c["tz"] == "Asia/Tokyo" for c in rpc.contexts)
    assert {c["lang"] for c in rpc.contexts[1:]} == {"en_US", "ja_JP"}
    assert {method for _, method in rpc.calls} <= {"search_read", "read", "get_views"}


def test_find_and_filters_match_both_languages_and_normalize_fullwidth(bilingual):
    assert [key for key, _, _ in sitemap.find(bilingual, 'sale.menu_orders')] == ['menu:2']
    assert sitemap.find(bilingual, "販売 注文")[0][0] == "menu:2"
    assert sitemap.find(bilingual, "ＳＡＬＥＳ ＯＲＤＥＲＳ")[0][0] == "menu:2"
    assert sitemap.find(bilingual, "確認")[0][0] == "view:sale.order:form:2"
    assert sitemap.find(bilingual, "顧客")[0][0] == "view:sale.order:form:2"
    assert '"Confirm"' in sitemap.text_buttons(bilingual, "注文", match="確認", lang="en")
    assert '"Customer"' in sitemap.text_fields(bilingual, "Orders", match="顧客", lang="en")
    assert "販売注文" in sitemap.text_fields(bilingual, "Orders", view="list", match="Sales Order", lang="ja")


def test_japanese_responses_cover_labels_fields_choices_paths_and_errors(bilingual):
    before = deepcopy(bilingual)
    assert "サイトマップ" in sitemap.text_find(bilingual, "Orders", lang="ja")
    buttons = sitemap.text_buttons(bilingual, "Orders", match="Confirm", lang="ja")
    assert "ボタン「確認」" in buttons and "実行するまで" in buttons
    assert "button[name='action_confirm']" in buttons and "state != 'draft'" in buttons
    fields = sitemap.text_fields(bilingual, "Orders", required=True, lang="ja")
    assert "顧客" in fields and "必須" in fields and "res.partner" in fields
    choices = sitemap.text_fields(bilingual, "Orders", view="list", lang="ja")
    assert "draft=見積, sale=販売注文" in choices and "th[data-name='state']" in choices
    route = sitemap.text_path(bilingual, "", "Cancel order", lang="ja")
    assert "開始: メニュー 販売" in route and "ボタン「取消」をクリック" in route and "ダイアログ" in route
    assert "見つかりません" in sitemap.text_path(bilingual, "window:11", "menu:2", lang="ja")
    with pytest.raises(sitemap.SitemapError, match="一致する項目がありません"):
        sitemap.text_path(bilingual, "", "nonexistent", lang="ja")
    assert "一致する項目がありません" in sitemap.text_fields(bilingual, "nonexistent", lang="ja")
    assert "一致する項目がありません" in sitemap.text_buttons(bilingual, "nonexistent", lang="ja")
    assert "一致する項目がありません" in sitemap.text_find(bilingual, "nonexistent", lang="ja")
    assert "読み取れませんでした" in sitemap.unresolved_text(bilingual["unresolved"][0], "ja")
    assert bilingual == before


def test_language_availability_is_reported_without_fake_translations():
    data = sitemap.build(BilingualRpc(active=("en_US",)))
    assert data["meta"]["languages"] == ["en_US"] and data["meta"]["missing_languages"] == ["ja_JP"]
    assert "未有効の言語: ja_JP" in sitemap.header(data, "ja")
    assert "ja_JP" not in data["nodes"]["menu:2"]["labels"]
    with pytest.raises(sitemap.SitemapError, match="ja_JP"):
        sitemap.build(BilingualRpc(active=("en_US",)), languages="en,ja", lang="ja")
    with pytest.raises(sitemap.SitemapError, match="有効にしてください"):
        sitemap.build(BilingualRpc(active=()), lang="ja")
    with pytest.raises(sitemap.SitemapError, match="en_US / ja_JP"):
        sitemap.build(BilingualRpc(), languages="fr_FR")
    japanese = sitemap.build(BilingualRpc(), languages="ja,ja_JP")
    assert japanese["meta"]["languages"] == ["ja_JP"] and "注文" in japanese["nodes"]["menu:2"]["label"]


def test_bilingual_sitemap_round_trips_and_legacy_maps_still_work(bilingual, ws):
    file = sitemap.save("demo", bilingual)
    assert "顧客" in file.read_text(encoding="utf-8")
    assert "顧客" in sitemap.text_fields(sitemap.load("demo"), "Orders", lang="ja")
    legacy = deepcopy(bilingual)
    # Old schema-1 maps have no per-language labels or language metadata.
    for node in legacy["nodes"].values():
        node.pop("labels", None)
        for item in node.get("fields", []) + node.get("buttons", []):
            item.pop("labels", None)
            item.pop("options_by_language", None)
    for edge in legacy["edges"]:
        edge.pop("labels", None)
    for key in ("languages", "default_language", "missing_languages"):
        legacy["meta"].pop(key, None)
    assert "ボタン「Confirm」" in sitemap.text_buttons(legacy, "Orders", lang="ja")


def test_cli_and_mcp_use_workspace_language_and_allow_overrides(bilingual, ws, store, capsys):
    ws["language"] = "ja"
    workspace.save(ws, store)
    sitemap.save("demo", bilingual)
    assert cli.main(["sitemap", "buttons", "Orders", "--match", "Confirm", "--ws", "demo"]) == 0
    assert "ボタン「確認」" in capsys.readouterr().out
    assert cli.main(["sitemap", "fields", "注文", "--match", "顧客", "--lang", "en", "--ws", "demo"]) == 0
    assert '"Customer"' in capsys.readouterr().out
    assert cli.main(["sitemap", "show", "--ws", "demo"]) == 0
    assert "サイトマップ" in capsys.readouterr().out
    assert "ボタン「取消」をクリック" in mcp_server.call("demo", "sitemap_path", {"target": "Cancel order"})
    assert "click the button" in mcp_server.call("demo", "sitemap_path", {"target": "注文取消", "lang": "en"})
    assert "顧客" in mcp_server.call("demo", "sitemap_fields", {"screen": "Orders", "match": "Customer"})
    assert "ボタン" in mcp_server.call("demo", "sitemap_buttons", {"screen": "Orders", "match": "確認"})
    assert "メニュー" in mcp_server.call("demo", "sitemap_find", {"text": "Orders"})
    assert all("lang" in tool["inputSchema"]["properties"] for tool in mcp_server.TOOLS)


def test_cli_build_forwards_languages_and_localizes_summary(ws, store, monkeypatch, capsys):
    rpc = BilingualRpc()

    class Page:
        def evaluate(self, script):
            return "19.0"

    class Odoo:
        def __init__(self, *args):
            self.rpc = rpc

    @contextmanager
    def signed_in(*args, **kwargs):
        yield Page()

    monkeypatch.setattr(cli, "_session_and_workspace", lambda args: (None, ws))
    monkeypatch.setattr(cli, "default_store", lambda: store)
    monkeypatch.setattr(cli.probe, "signed_in", signed_in)
    monkeypatch.setattr("oqa.odoo_ui.Odoo", Odoo)
    assert cli.main(["sitemap", "build", "--languages", "en,ja", "--lang", "ja", "--ws", "demo"]) == 0
    assert sitemap.load("demo")["meta"]["languages"] == ["en_US", "ja_JP"]
    output = capsys.readouterr().out
    assert "保存先" in output and "読み取りできなかった項目" in output and "読み取れませんでした" in output
