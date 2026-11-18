"""Database evidence through pgAdmin: the SQL guard, what the analysis must say about it, and that s.db counts as the check and the frame of its step."""
import copy

import pytest

from conftest import filled
from oqa import analysis, scaffold
from oqa.pgadmin import PgAdminError, check_expect, check_select
from oqa.store import SchemaError

GOOD = {"sql": "select id, name from res_partner where id = 7", "expect": [{"id": "7", "name": "Azure"}], "mark": ["name"], "why": "the partner exists"}


@pytest.mark.parametrize("sql", ["select * from res_partner", "  WITH a AS (select 1) select * from a;", "select comment, lock, create_date, \"drop\" from t -- c",
                                 "select 'delete from x' as s", "show server_version"])
def test_a_read_only_query_passes(sql):
    check_select(sql)


@pytest.mark.parametrize("sql", ["", "delete from a", "select 1; select 2", "update res_partner set name='a'", "drop table a", "select * into t2 from t", "insert into a values (1)",
                                 "with d as (delete from a returning *) select * from d", "create table x(a int)", "truncate a", "explain analyze delete from a", "copy a to stdout"])
def test_anything_that_writes_is_refused(sql):
    with pytest.raises(PgAdminError):
        check_select(sql)


def test_expected_rows_are_a_list_of_column_value_pairs():
    assert check_expect([{"id": 7}]) == [{"id": "7"}]
    for bad in (None, [], [{}], ["x"], {"id": 1}):
        with pytest.raises(PgAdminError):
            check_expect(bad)


def with_pg(config):
    config = copy.deepcopy(config)
    config["pgadmin"] = {"group": "Servers", "server": "localhost", "database": "db_x"}
    return config


def test_with_pgadmin_configured_every_case_must_answer_the_database_question(session, ws):
    config = with_pg(ws)
    cases = session.cases()
    data = filled(cases, config)
    assert all(c["db_evidence"] == {"needed": None, "reason": ""} for c in data["cases"])          # the template asks
    with pytest.raises(SchemaError, match="db_evidence"):
        analysis.validate(data, cases, config)
    for c in data["cases"]:
        c["db_evidence"] = {"needed": False, "reason": "the case only looks at screens"}
    analysis.validate(data, cases, config)


def test_a_case_that_needs_database_evidence_has_a_step_with_the_query_and_rows(session, ws):
    config = with_pg(ws)
    cases = session.cases()
    data = filled(cases, config)
    for c in data["cases"]:
        c["db_evidence"] = {"needed": False, "reason": "screens only"}
    first = data["cases"][0]
    first["db_evidence"] = {"needed": True, "reason": "the order must exist in the table"}
    with pytest.raises(SchemaError, match="needs `db`"):
        analysis.validate(data, cases, config)
    first["steps"][-1]["db"] = dict(GOOD)
    analysis.validate(data, cases, config)
    first["steps"][-1]["db"]["sql"] = "delete from res_partner"
    with pytest.raises(SchemaError, match="only a SELECT|writes"):
        analysis.validate(data, cases, config)
    first["steps"][-1]["db"] = dict(GOOD, expect=[])
    with pytest.raises(SchemaError, match="expect"):
        analysis.validate(data, cases, config)
    first["steps"][-1]["db"] = dict(GOOD)
    first["db_evidence"]["needed"] = False
    with pytest.raises(SchemaError, match="needed is false"):
        analysis.validate(data, cases, config)


def test_a_db_step_without_pgadmin_in_the_workspace_is_refused(session, ws):
    cases = session.cases()
    data = filled(cases, ws)
    data["cases"][0]["steps"][0]["db"] = dict(GOOD)
    with pytest.raises(SchemaError, match="no pgAdmin settings"):
        analysis.validate(data, cases, ws)
    assert "db_evidence" not in data["cases"][1]                         # no pgAdmin, no question: the feature is optional


def test_the_script_skeleton_and_checks_know_s_db(tmp_path):
    item = {"profile": "Admin", "steps": [{"n": 1, "action": "Save", "expected": "ok", "db": dict(GOOD)}]}
    text = scaffold.render({"id": "C-1", "title": "t"}, item, "sha")
    assert "oqa try C-1 --step 1 --db" in text and "select id, name from res_partner where id = 7" in text
    script = tmp_path / "c.py"
    script.write_text("def step_1(t, s):\n    o = t.o\n    s.db('select 1', expect=[{'id': '1'}])\n", encoding="utf-8")
    assert scaffold.unchecked_steps(script) == [] and scaffold.unhighlighted_steps(script) == []          # s.db is the check and the frame
