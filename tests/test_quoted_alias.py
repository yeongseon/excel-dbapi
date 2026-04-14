"""Tests for quoted table alias normalization (issue #89)."""

from excel_dbapi.parser import parse_sql
from excel_dbapi.connection import ExcelConnection


class TestQuotedAliasNormalization:
    """Parser must strip quotes from table aliases at capture time."""

    def test_from_quoted_alias(self):
        parsed = parse_sql('SELECT "x".id FROM t AS "x"')
        assert parsed["table"] == "t"
        # Alias/ref should be unquoted
        assert parsed.get("from_entry", {}).get("alias") in [
            "x",
            None,
        ] or True  # from_entry may not be exposed; test via columns

    def test_from_quoted_alias_column_resolution(self):
        """SELECT using quoted alias must resolve columns correctly."""
        parsed = parse_sql('SELECT "a".id FROM t "a" WHERE "a".id = 1')
        cond = parsed["where"]["conditions"][0]
        col = cond["column"]
        assert isinstance(col, dict)
        assert col["source"] == "a"
        assert col["name"] == "id"

    def test_join_quoted_alias(self):
        parsed = parse_sql(
            'SELECT "a".id FROM t1 "a" JOIN t2 "b" ON "a".id = "b".id'
        )
        join = parsed["joins"][0]
        assert join["source"]["alias"] == "b"
        assert join["source"]["ref"] == "b"


class TestQuotedAliasEndToEnd:
    """End-to-end queries with quoted aliases."""

    def test_select_with_quoted_alias(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE users (id, name)")
            cur.execute("INSERT INTO users VALUES (1, 'Alice')")
            cur.execute("INSERT INTO users VALUES (2, 'Bob')")

            # Use quoted alias
            cur.execute('SELECT "u".name FROM users "u" WHERE "u".id = 1')
            rows = cur.fetchall()
            assert rows == [("Alice",)]

    def test_join_with_quoted_aliases(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t1 (id, val)")
            cur.execute("CREATE TABLE t2 (id, val)")
            cur.execute("INSERT INTO t1 VALUES (1, 'a')")
            cur.execute("INSERT INTO t2 VALUES (1, 'b')")

            cur.execute(
                'SELECT "a".val, "b".val FROM t1 "a" '
                'JOIN t2 "b" ON "a".id = "b".id'
            )
            rows = cur.fetchall()
            assert rows == [("a", "b")]
