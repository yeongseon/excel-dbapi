"""Tests for _tokenize splitting comparison operators without spaces (issue #86)."""


from excel_dbapi.parser.tokenizer import _tokenize
from excel_dbapi.parser import parse_sql
from excel_dbapi.connection import ExcelConnection


class TestTokenizeOperatorSplitting:
    """_tokenize must emit comparison operators as standalone tokens."""

    def test_equals_no_spaces(self):
        assert _tokenize("id=1") == ["id", "=", "1"]

    def test_not_equals_no_spaces(self):
        assert _tokenize("id!=1") == ["id", "!=", "1"]

    def test_diamond_no_spaces(self):
        assert _tokenize("id<>1") == ["id", "<>", "1"]

    def test_less_than_no_spaces(self):
        assert _tokenize("x<5") == ["x", "<", "5"]

    def test_less_than_or_equal_no_spaces(self):
        assert _tokenize("x<=5") == ["x", "<=", "5"]

    def test_greater_than_no_spaces(self):
        assert _tokenize("x>5") == ["x", ">", "5"]

    def test_greater_than_or_equal_no_spaces(self):
        assert _tokenize("x>=5") == ["x", ">=", "5"]

    def test_equals_with_spaces(self):
        """Existing space-separated operators must still work."""
        assert _tokenize("id = 1") == ["id", "=", "1"]

    def test_not_equals_with_spaces(self):
        assert _tokenize("id != 1") == ["id", "!=", "1"]

    def test_complex_expression_no_spaces(self):
        """Multiple operators in one expression."""
        tokens = _tokenize("a=1 AND b!=2 AND c<>3")
        assert tokens == ["a", "=", "1", "AND", "b", "!=", "2", "AND", "c", "<>", "3"]

    def test_operator_inside_single_quotes(self):
        """Operators inside string literals must NOT be split."""
        tokens = _tokenize("name = 'a=b'")
        assert tokens == ["name", "=", "'a=b'"]

    def test_operator_inside_double_quotes(self):
        """Operators inside double-quoted identifiers must NOT be split."""
        tokens = _tokenize('"col=1" = 5')
        assert tokens == ['"col=1"', "=", "5"]

    def test_less_than_in_string_literal(self):
        tokens = _tokenize("x = '<test>'")
        assert tokens == ["x", "=", "'<test>'"]

    def test_parenthesized_with_operators(self):
        tokens = _tokenize("(x=1)")
        assert tokens == ["(", "x", "=", "1", ")"]

    def test_chained_comparisons_no_spaces(self):
        tokens = _tokenize("a>=1 AND b<=2")
        assert tokens == ["a", ">=", "1", "AND", "b", "<=", "2"]


class TestParserWithOperatorsNoSpaces:
    """parse_sql must handle WHERE clauses with no spaces around operators."""

    def test_select_where_equals_no_space(self):
        parsed = parse_sql("SELECT * FROM t WHERE id=1")
        assert parsed["where"]["conditions"][0]["column"] == "id"
        assert parsed["where"]["conditions"][0]["operator"] == "="
        assert parsed["where"]["conditions"][0]["value"] == 1

    def test_select_where_not_equals_no_space(self):
        parsed = parse_sql("SELECT * FROM t WHERE status!='done'")
        assert parsed["where"]["conditions"][0]["column"] == "status"
        assert parsed["where"]["conditions"][0]["operator"] == "!="
        assert parsed["where"]["conditions"][0]["value"] == "done"

    def test_select_where_less_than_no_space(self):
        parsed = parse_sql("SELECT * FROM t WHERE age<30")
        cond = parsed["where"]["conditions"][0]
        assert cond["column"] == "age"
        assert cond["operator"] == "<"
        assert cond["value"] == 30

    def test_select_where_greater_equal_no_space(self):
        parsed = parse_sql("SELECT * FROM t WHERE score>=90")
        cond = parsed["where"]["conditions"][0]
        assert cond["column"] == "score"
        assert cond["operator"] == ">="
        assert cond["value"] == 90


class TestEndToEndOperatorsNoSpaces:
    """End-to-end: run SQL with operators without spaces against real data."""

    def test_select_where_equals_no_space(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (id, name)")
            cur.execute("INSERT INTO t VALUES (1, 'Alice')")
            cur.execute("INSERT INTO t VALUES (2, 'Bob')")

            cur.execute("SELECT * FROM t WHERE id=1")
            rows = cur.fetchall()
            assert rows == [(1, "Alice")]

    def test_select_where_not_equals_no_space(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (id, name)")
            cur.execute("INSERT INTO t VALUES (1, 'Alice')")
            cur.execute("INSERT INTO t VALUES (2, 'Bob')")

            cur.execute("SELECT * FROM t WHERE id!=1")
            rows = cur.fetchall()
            assert rows == [(2, "Bob")]

    def test_select_where_comparison_no_space(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (id, score)")
            cur.execute("INSERT INTO t VALUES (1, 80)")
            cur.execute("INSERT INTO t VALUES (2, 90)")
            cur.execute("INSERT INTO t VALUES (3, 70)")

            cur.execute("SELECT id FROM t WHERE score>=80 ORDER BY id")
            rows = cur.fetchall()
            assert rows == [(1,), (2,)]

    def test_update_where_no_space(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (id, name)")
            cur.execute("INSERT INTO t VALUES (1, 'Alice')")
            cur.execute("INSERT INTO t VALUES (2, 'Bob')")

            cur.execute("UPDATE t SET name='Ann' WHERE id=1")
            cur.execute("SELECT name FROM t WHERE id=1")
            assert cur.fetchone() == ("Ann",)

    def test_delete_where_no_space(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (id, name)")
            cur.execute("INSERT INTO t VALUES (1, 'Alice')")
            cur.execute("INSERT INTO t VALUES (2, 'Bob')")

            cur.execute("DELETE FROM t WHERE id<>1")
            cur.execute("SELECT * FROM t")
            assert cur.fetchall() == [(1, "Alice")]
