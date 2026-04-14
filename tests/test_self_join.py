"""Tests for self-join with distinct aliases (issue #87)."""

from excel_dbapi.parser import parse_sql
from excel_dbapi.connection import ExcelConnection


class TestSelfJoinParsing:
    """Self-joins with distinct aliases must be accepted by the parser."""

    def test_self_join_with_aliases(self):
        parsed = parse_sql("SELECT u.id, v.id FROM users u JOIN users v ON u.id = v.id")
        assert parsed["table"] == "users"
        assert len(parsed["joins"]) == 1
        assert parsed["joins"][0]["source"]["table"] == "users"
        assert parsed["joins"][0]["source"]["alias"] == "v"

    def test_self_join_with_as_aliases(self):
        parsed = parse_sql(
            "SELECT a.id FROM orders AS a JOIN orders AS b ON a.id = b.id"
        )
        assert parsed["table"] == "orders"
        assert parsed["joins"][0]["source"]["table"] == "orders"
        assert parsed["joins"][0]["source"]["alias"] == "b"

    def test_self_join_left(self):
        parsed = parse_sql(
            "SELECT a.id FROM t AS a LEFT JOIN t AS b ON a.id = b.id"
        )
        assert parsed["joins"][0]["type"] == "LEFT"
        assert parsed["joins"][0]["source"]["alias"] == "b"

    def test_different_tables_no_alias_still_works(self):
        parsed = parse_sql("SELECT * FROM t1 JOIN t2 ON t1.id = t2.id")
        assert parsed["table"] == "t1"
        assert parsed["joins"][0]["source"]["table"] == "t2"

    def test_same_alias_rejected(self):
        """Two sources with the SAME alias should still be rejected."""
        import pytest

        with pytest.raises(ValueError, match="Ambiguous"):
            parse_sql("SELECT * FROM t AS x JOIN t2 AS x ON x.id = x.id")


class TestSelfJoinEndToEnd:
    """End-to-end self-join queries against real data."""

    def test_self_join_select(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE emp (id, name, manager_id)")
            cur.execute("INSERT INTO emp VALUES (1, 'Alice', NULL)")
            cur.execute("INSERT INTO emp VALUES (2, 'Bob', 1)")
            cur.execute("INSERT INTO emp VALUES (3, 'Carol', 1)")

            cur.execute(
                "SELECT e.name, m.name FROM emp e "
                "JOIN emp m ON e.manager_id = m.id"
            )
            rows = cur.fetchall()
            assert sorted(rows) == [("Bob", "Alice"), ("Carol", "Alice")]

    def test_self_join_left(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE emp (id, name, manager_id)")
            cur.execute("INSERT INTO emp VALUES (1, 'Alice', NULL)")
            cur.execute("INSERT INTO emp VALUES (2, 'Bob', 1)")

            cur.execute(
                "SELECT e.name, m.name FROM emp e "
                "LEFT JOIN emp m ON e.manager_id = m.id "
                "ORDER BY e.id"
            )
            rows = cur.fetchall()
            assert rows == [("Alice", None), ("Bob", "Alice")]
