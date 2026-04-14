"""Tests for quoted identifiers in DDL (CREATE/ALTER TABLE) — issue #88."""

from excel_dbapi.parser import parse_sql
from excel_dbapi.connection import ExcelConnection


class TestCreateTableQuotedIdentifiers:
    """CREATE TABLE with double-quoted column names."""

    def test_quoted_column_name(self):
        parsed = parse_sql('CREATE TABLE t ("Full Name")')
        assert parsed["columns"] == ["Full Name"]

    def test_quoted_column_with_type(self):
        parsed = parse_sql('CREATE TABLE t ("Full Name" TEXT)')
        assert parsed["columns"] == ["Full Name"]
        assert parsed["column_definitions"][0]["name"] == "Full Name"
        assert parsed["column_definitions"][0]["type_name"] == "TEXT"

    def test_mixed_quoted_and_bare(self):
        parsed = parse_sql('CREATE TABLE t (id, "Full Name" TEXT, age INTEGER)')
        assert parsed["columns"] == ["id", "Full Name", "age"]
        assert parsed["column_definitions"][0]["type_name"] == "TEXT"
        assert parsed["column_definitions"][1]["name"] == "Full Name"
        assert parsed["column_definitions"][1]["type_name"] == "TEXT"
        assert parsed["column_definitions"][2]["type_name"] == "INTEGER"

    def test_quoted_column_with_spaces_no_type(self):
        parsed = parse_sql('CREATE TABLE t ("First Name", "Last Name")')
        assert parsed["columns"] == ["First Name", "Last Name"]

    def test_quoted_column_with_escaped_quotes(self):
        parsed = parse_sql('CREATE TABLE t ("say ""hello""")')
        assert parsed["columns"] == ['say "hello"']


class TestAlterTableQuotedIdentifiers:
    """ALTER TABLE with double-quoted column names."""

    def test_add_quoted_column(self):
        parsed = parse_sql('ALTER TABLE t ADD COLUMN "Full Name" TEXT')
        assert parsed["column"] == "Full Name"

    def test_drop_quoted_column(self):
        parsed = parse_sql('ALTER TABLE t DROP COLUMN "Full Name"')
        assert parsed["column"] == "Full Name"

    def test_rename_quoted_columns(self):
        parsed = parse_sql(
            'ALTER TABLE t RENAME COLUMN "Old Name" TO "New Name"'
        )
        assert parsed["old_column"] == "Old Name"
        assert parsed["new_column"] == "New Name"


class TestDDLQuotedEndToEnd:
    """End-to-end DDL with quoted column identifiers."""

    def test_create_and_insert_quoted_columns(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute('CREATE TABLE t ("Full Name" TEXT, age INTEGER)')
            cur.execute(
                'INSERT INTO t ("Full Name", age) VALUES (\'Alice\', 30)'
            )
            cur.execute("SELECT * FROM t")
            rows = cur.fetchall()
            assert rows == [("Alice", 30)]

    def test_alter_add_quoted_column(self, tmp_path):
        file = tmp_path / "test.xlsx"
        with ExcelConnection(str(file), autocommit=True, create=True) as conn:
            cur = conn.cursor()
            cur.execute("CREATE TABLE t (id)")
            cur.execute('ALTER TABLE t ADD COLUMN "Full Name" TEXT')
            cur.execute("INSERT INTO t (id, \"Full Name\") VALUES (1, 'Alice')")
            cur.execute("SELECT * FROM t")
            rows = cur.fetchall()
            assert rows == [(1, "Alice")]
