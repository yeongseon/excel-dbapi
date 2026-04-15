import pytest
from excel_dbapi.exceptions import DatabaseError

from excel_dbapi.parser import parse_sql


def test_scalar_subquery_in_where_is_accepted():
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = (SELECT id FROM Sheet2)")
    assert parsed is not None


def test_rejects_for_update():
    with pytest.raises(DatabaseError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users FOR UPDATE")


def test_rejects_nulls_last():
    with pytest.raises(DatabaseError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users ORDER BY id ASC NULLS LAST")


def test_accepts_cross_join_without_on():
    parsed = parse_sql("SELECT a.id FROM t1 a CROSS JOIN t2 b")
    assert parsed["joins"][0]["type"] == "CROSS"
    assert parsed["joins"][0]["on"] is None


def test_accepts_full_outer_join():
    parsed = parse_sql("SELECT a.id FROM t1 a FULL OUTER JOIN t2 b ON a.id = b.id")
    assert parsed["joins"][0]["type"] == "FULL"


def test_allows_join_with_select_star():
    parsed = parse_sql("SELECT * FROM t1 a JOIN t2 b ON a.id = b.id")
    assert parsed["columns"] == ["*"]
