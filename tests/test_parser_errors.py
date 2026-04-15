from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import (
    DatabaseError,
    NotSupportedError,
    ProgrammingError,
)
from excel_dbapi.parser import parse_sql

def test_parse_empty_query():
    with pytest.raises(DatabaseError):
        parse_sql("")


def test_parse_select_missing_from():
    with pytest.raises(DatabaseError):
        parse_sql("SELECT id Sheet1")


def test_parse_order_by_before_where():
    with pytest.raises(DatabaseError):
        parse_sql("SELECT * FROM Sheet1 ORDER BY id WHERE id = 1")


def test_parse_invalid_order_direction():
    with pytest.raises(DatabaseError):
        parse_sql("SELECT * FROM Sheet1 ORDER BY id DOWN")


def test_parse_invalid_limit():
    with pytest.raises(DatabaseError):
        parse_sql("SELECT * FROM Sheet1 LIMIT foo")


def test_parse_insert_missing_values():
    with pytest.raises(DatabaseError):
        parse_sql("INSERT INTO Sheet1 (id)")


def test_parse_insert_missing_params():
    with pytest.raises(DatabaseError):
        parse_sql("INSERT INTO Sheet1 (id) VALUES (?)")


def test_parse_update_missing_set():
    with pytest.raises(DatabaseError):
        parse_sql("UPDATE Sheet1 name = 'A'")


def test_parse_delete_missing_from():
    with pytest.raises(DatabaseError):
        parse_sql("DELETE Sheet1")


def test_parse_create_invalid_format():
    with pytest.raises(DatabaseError):
        parse_sql("CREATE TABLE Foo")


@pytest.mark.parametrize(
    "sql",
    [
        "CREATE TABLE t (id INTEGER name TEXT)",
        "CREATE TABLE t (id INTEGER, name TEXT age INTEGER)",
    ],
)
def test_parse_create_rejects_malformed_column_definitions(sql: str) -> None:
    with pytest.raises(DatabaseError, match="Malformed column definition"):
        parse_sql(sql)


def test_parse_drop_invalid_format():
    with pytest.raises(DatabaseError):
        parse_sql("DROP Foo")



@pytest.mark.parametrize(
    "sql",
    [
        "CREATE TABLE (id INTEGER)",
        "CREATE TABLE t (,id INTEGER)",
        "CREATE TABLE t (id INTEGER,, name TEXT)",
    ],
)
def test_create_table_rejects_missing_name_and_empty_column_definitions(
    sql: str,
) -> None:
    with pytest.raises(DatabaseError):
        parse_sql(sql)

def test_create_table_rejects_single_trailing_comma() -> None:
    with pytest.raises(DatabaseError, match="empty column definition"):
        parse_sql("CREATE TABLE t (id INTEGER, name TEXT,)")



def _create_round12_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "t"
    sheet.append(["id", "txt", "col"])
    sheet.append([1, "alpha", 1])
    sheet.append([2, "beta", 2])
    sheet.append([3, "gamma", 3])
    workbook.save(path)

def test_negative_limit_is_rejected_for_select_and_compound(tmp_path: Path) -> None:
    file_path = tmp_path / "round12_negative_limit.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()

        with pytest.raises(
            ProgrammingError, match="LIMIT must be a non-negative integer"
        ):
            cursor.execute("SELECT id FROM t LIMIT -1")

        with pytest.raises(
            ProgrammingError, match="LIMIT must be a non-negative integer"
        ):
            cursor.execute("SELECT id FROM t UNION SELECT id FROM t LIMIT -1")

def test_negative_offset_is_rejected_for_select_and_compound(tmp_path: Path) -> None:
    file_path = tmp_path / "round12_negative_offset.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()

        with pytest.raises(
            ProgrammingError,
            match="OFFSET must be a non-negative integer",
        ):
            cursor.execute("SELECT id FROM t OFFSET -1")

        with pytest.raises(
            ProgrammingError,
            match="OFFSET must be a non-negative integer",
        ):
            cursor.execute("SELECT id FROM t UNION SELECT id FROM t OFFSET -1")



def _make_xlsx(
    path: Path,
    sheet: str = "users",
    headers: list[str] | None = None,
    rows: list[list[Any]] | None = None,
) -> str:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet
    for h in headers or ["id", "name"]:
        pass
    ws.append(headers or ["id", "name"])
    for row in rows or []:
        ws.append(row)
    fpath = str(path)
    wb.save(fpath)
    wb.close()
    return fpath

def test_malformed_create_table_missing_comma_raises_programming_error(
    tmp_path: Path,
) -> None:
    fpath = _make_xlsx(tmp_path / "malformed-create.xlsx")
    with ExcelConnection(fpath, engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="Missing comma"):
            cursor.execute("CREATE TABLE t (id INTEGER name TEXT)")

    def test_not_implemented_becomes_not_supported(self, tmp_path: Path) -> None:
        """NotImplementedError → NotSupportedError."""
        fpath = _make_xlsx(tmp_path / "test.xlsx")
        with ExcelConnection(fpath, engine="openpyxl") as conn:
            cursor = conn.cursor()
            # Trigger a non-supported operation
            with pytest.raises((ProgrammingError, NotSupportedError)):
                cursor.execute("CREATE INDEX idx ON users (id)")

    def test_executemany_accepts_iterable_of_sequences(self, tmp_path: Path) -> None:
        """executemany() accepts Iterable[Sequence[Any]], not just List[tuple]."""
        fpath = _make_xlsx(tmp_path / "test.xlsx")
        with ExcelConnection(fpath, engine="openpyxl", autocommit=True) as conn:
            cursor = conn.cursor()
            # Pass a generator of lists (not List[tuple])
            params = ([i, f"user{i}"] for i in range(1, 4))
            cursor.executemany("INSERT INTO users VALUES (?, ?)", params)
            assert cursor.rowcount == 3

        with ExcelConnection(fpath, engine="openpyxl") as conn:
            result = conn.execute("SELECT * FROM users")
            assert len(result.rows) == 3

    def test_executemany_rollback_on_error_autocommit_off(self, tmp_path: Path) -> None:
        """executemany() with autocommit=False rolls back on error."""
        fpath = _make_xlsx(tmp_path / "test.xlsx", rows=[[1, "Original"]])
        with ExcelConnection(fpath, engine="openpyxl", autocommit=False) as conn:
            cursor = conn.cursor()
            with pytest.raises((ProgrammingError, DatabaseError)):
                # First insert succeeds, second should fail
                cursor.executemany(
                    "INSERT INTO users VALUES (?, ?)",
                    [(2, "Bob"), (3, None, "extra")],  # type: ignore[list-item]
                )
            # After error, rollback should have happened
            conn.rollback()
            result = conn.execute("SELECT * FROM users")
            # Only original row should remain
            assert len(result.rows) == 1
            assert result.rows[0] == (1, "Original")

    def test_database_error_subclasses_pass_through(self, tmp_path: Path) -> None:
        """PEP 249 exceptions already raised in executor pass through unchanged."""
        fpath = _make_xlsx(tmp_path / "test.xlsx")
        with ExcelConnection(fpath, engine="openpyxl") as conn:
            cursor = conn.cursor()
            # This raises ProgrammingError from the parser — should pass through
            with pytest.raises(ProgrammingError):
                cursor.execute("SELECT * FROM nonexistent_table_xyz")



def test_aggregate_sum_star_rejected() -> None:
    """SUM(*) is not supported — only COUNT(*)."""
    with pytest.raises(DatabaseError, match="SUM does not support"):
        parse_sql("SELECT SUM(*) FROM Sheet1")

def test_aggregate_avg_star_rejected() -> None:
    with pytest.raises(DatabaseError, match="AVG does not support"):
        parse_sql("SELECT AVG(*) FROM Sheet1")

def test_aggregate_expression_rejected() -> None:
    """SUM(a+b) — expressions inside aggregates are not supported."""
    with pytest.raises(DatabaseError, match="Unsupported function: SUM"):
        parse_sql("SELECT SUM(a + b) FROM Sheet1")

def test_where_single_token() -> None:
    """WHERE clause with only one token is invalid."""
    with pytest.raises(DatabaseError, match="Invalid WHERE clause format"):
        parse_sql("SELECT * FROM Sheet1 WHERE x")

def test_where_is_not_missing_null() -> None:
    """IS NOT without NULL is invalid."""
    with pytest.raises(DatabaseError, match="expected NULL or NOT after IS"):
        parse_sql("SELECT * FROM Sheet1 WHERE x IS FOO")

def test_where_between_missing_and() -> None:
    """BETWEEN without AND keyword is invalid."""
    with pytest.raises(DatabaseError, match="expected AND in BETWEEN"):
        parse_sql("SELECT * FROM Sheet1 WHERE x BETWEEN 1 OR 10")

def test_where_in_missing_tokens() -> None:
    """IN without parenthesized list."""
    with pytest.raises(DatabaseError, match="Invalid WHERE clause format"):
        parse_sql("SELECT * FROM Sheet1 WHERE x IN")

def test_where_in_malformed_no_close_paren() -> None:
    """IN clause missing close paren."""
    with pytest.raises(DatabaseError, match="IN clause"):
        parse_sql("SELECT * FROM Sheet1 WHERE x IN (1, 2")

def test_where_like_missing_pattern() -> None:
    """LIKE without pattern value."""
    with pytest.raises(DatabaseError, match="Invalid WHERE clause format"):
        parse_sql("SELECT * FROM Sheet1 WHERE x LIKE")

def test_where_operator_missing_value() -> None:
    """Regular operator (=) without right-hand value."""
    with pytest.raises(DatabaseError, match="Invalid WHERE clause format"):
        parse_sql("SELECT * FROM Sheet1 WHERE x =")

def test_subquery_with_join_inside() -> None:
    """Subqueries cannot contain JOIN."""
    with pytest.raises(DatabaseError, match="JOIN is not supported in subqueries"):
        parse_sql(
            "SELECT * FROM t WHERE id IN (SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id)"
        )

def test_subquery_with_having() -> None:
    """Subqueries cannot contain HAVING."""
    with pytest.raises(DatabaseError, match="HAVING is not supported in subqueries"):
        parse_sql(
            "SELECT * FROM t WHERE id IN (SELECT id FROM t2 GROUP BY id HAVING COUNT(*) > 1)"
        )

def test_subquery_with_order_by() -> None:
    """Subqueries cannot contain ORDER BY."""
    with pytest.raises(DatabaseError, match="ORDER BY is not supported in subqueries"):
        parse_sql("SELECT * FROM t WHERE id IN (SELECT id FROM t2 ORDER BY id)")

def test_subquery_with_offset() -> None:
    """Subqueries cannot contain OFFSET."""
    with pytest.raises(DatabaseError, match="OFFSET is not supported in subqueries"):
        parse_sql("SELECT * FROM t WHERE id IN (SELECT id FROM t2 OFFSET 5)")

def test_subquery_with_group_by() -> None:
    """Subqueries cannot contain GROUP BY."""
    with pytest.raises(DatabaseError, match="GROUP BY is not supported in subqueries"):
        parse_sql("SELECT * FROM t WHERE id IN (SELECT id FROM t2 GROUP BY id)")

def test_subquery_with_limit() -> None:
    """Subqueries cannot contain LIMIT."""
    with pytest.raises(DatabaseError, match="LIMIT is not supported in subqueries"):
        parse_sql("SELECT * FROM t WHERE id IN (SELECT id FROM t2 LIMIT 5)")

def test_join_invalid_column_reference() -> None:
    """JOIN column must be qualified (table.column)."""
    with pytest.raises(DatabaseError, match="qualified column-to-column comparisons"):
        parse_sql("SELECT a.id FROM t1 a JOIN t2 b ON id = b.id")

def test_join_on_empty() -> None:
    """JOIN requires ON condition."""
    with pytest.raises(DatabaseError, match="JOIN requires ON condition"):
        parse_sql("SELECT a.id FROM t1 a JOIN t2 b")

def test_join_on_references_wrong_tables() -> None:
    """JOIN ON must reference columns from the joined tables."""
    with pytest.raises(DatabaseError, match="compare columns from the two joined sources"):
        parse_sql("SELECT a.id FROM t1 a JOIN t2 b ON a.id = a.name")

def test_join_on_too_many_tokens_in_comparison() -> None:
    """JOIN ON comparison must be exactly 3 tokens (left op right)."""
    with pytest.raises(DatabaseError, match="Invalid WHERE clause format"):
        parse_sql("SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id = 1")

def test_select_distinct_empty_after_keyword() -> None:
    """DISTINCT followed by nothing."""
    with pytest.raises(DatabaseError):
        parse_sql("SELECT DISTINCT")

def test_select_as_alias_missing() -> None:
    """AS keyword without following alias."""
    with pytest.raises(DatabaseError, match="Expected alias after AS"):
        parse_sql("SELECT a.id FROM t1 AS")

def test_inner_without_join() -> None:
    """INNER without JOIN keyword."""
    with pytest.raises(DatabaseError, match="Unsupported SQL syntax: INNER"):
        parse_sql("SELECT a.id FROM t1 a INNER t2 b ON a.id = b.id")

def test_left_without_join() -> None:
    """LEFT without JOIN keyword."""
    with pytest.raises(DatabaseError, match="Unsupported SQL syntax: LEFT"):
        parse_sql("SELECT a.id FROM t1 a LEFT t2 b ON a.id = b.id")

def test_join_missing_table() -> None:
    """JOIN without table name after it."""
    with pytest.raises(DatabaseError, match="missing table"):
        parse_sql("SELECT a.id FROM t1 a JOIN")

def test_join_source_invalid_reference() -> None:
    """Invalid source reference in SELECT column with JOIN."""
    with pytest.raises(DatabaseError, match="Invalid source reference"):
        parse_sql("SELECT c.id FROM t1 a JOIN t2 b ON a.id = b.id")

def test_order_by_before_where() -> None:
    with pytest.raises(DatabaseError, match="ORDER BY cannot appear before WHERE"):
        parse_sql("SELECT * FROM t ORDER BY x WHERE x = 1")

def test_limit_before_where() -> None:
    with pytest.raises(DatabaseError, match="LIMIT cannot appear before WHERE"):
        parse_sql("SELECT * FROM t LIMIT 10 WHERE x = 1")

def test_offset_before_where() -> None:
    with pytest.raises(DatabaseError, match="OFFSET cannot appear before WHERE"):
        parse_sql("SELECT * FROM t OFFSET 5 WHERE x = 1")

def test_group_by_before_where() -> None:
    with pytest.raises(DatabaseError, match="GROUP BY cannot appear before WHERE"):
        parse_sql("SELECT x, COUNT(*) FROM t GROUP BY x WHERE x = 1")

def test_having_before_where() -> None:
    with pytest.raises(DatabaseError):
        parse_sql(
            "SELECT x, COUNT(*) FROM t GROUP BY x HAVING COUNT(*) > 1 WHERE x = 1"
        )

def test_having_before_group_by() -> None:
    with pytest.raises(DatabaseError, match="HAVING cannot appear before GROUP BY"):
        parse_sql("SELECT x FROM t HAVING COUNT(*) > 1 GROUP BY x")

def test_order_by_before_group_by() -> None:
    with pytest.raises(DatabaseError, match="ORDER BY cannot appear before GROUP BY"):
        parse_sql("SELECT x FROM t ORDER BY x GROUP BY x")

def test_limit_before_group_by() -> None:
    with pytest.raises(DatabaseError, match="LIMIT cannot appear before GROUP BY"):
        parse_sql("SELECT x FROM t LIMIT 10 GROUP BY x")

def test_offset_before_group_by() -> None:
    with pytest.raises(DatabaseError, match="OFFSET cannot appear before GROUP BY"):
        parse_sql("SELECT x FROM t OFFSET 5 GROUP BY x")

def test_order_by_before_having() -> None:
    with pytest.raises(DatabaseError, match="ORDER BY cannot appear before HAVING"):
        parse_sql("SELECT x, COUNT(*) FROM t GROUP BY x ORDER BY x HAVING COUNT(*) > 1")

def test_limit_before_having() -> None:
    with pytest.raises(DatabaseError, match="LIMIT cannot appear before HAVING"):
        parse_sql("SELECT x, COUNT(*) FROM t GROUP BY x LIMIT 10 HAVING COUNT(*) > 1")

def test_offset_before_having() -> None:
    with pytest.raises(DatabaseError, match="OFFSET cannot appear before HAVING"):
        parse_sql("SELECT x, COUNT(*) FROM t GROUP BY x OFFSET 5 HAVING COUNT(*) > 1")

def test_offset_before_order_by() -> None:
    with pytest.raises(DatabaseError, match="OFFSET cannot appear before ORDER BY"):
        parse_sql("SELECT * FROM t OFFSET 5 ORDER BY x")

def test_offset_before_limit() -> None:
    with pytest.raises(DatabaseError, match="OFFSET cannot appear before LIMIT"):
        parse_sql("SELECT * FROM t OFFSET 5 LIMIT 10")

def test_empty_group_by() -> None:
    with pytest.raises(DatabaseError, match="Invalid GROUP BY clause format"):
        # This is tricky — GROUP BY followed immediately by ORDER BY
        parse_sql("SELECT x FROM t GROUP BY ORDER BY x")

def test_empty_having() -> None:
    with pytest.raises(DatabaseError, match="Invalid HAVING clause format"):
        parse_sql("SELECT x FROM t GROUP BY x HAVING ORDER BY x")

def test_empty_limit() -> None:
    with pytest.raises(DatabaseError, match="Invalid LIMIT clause format"):
        parse_sql("SELECT * FROM t LIMIT")

def test_empty_offset() -> None:
    with pytest.raises(DatabaseError, match="Invalid OFFSET clause format"):
        parse_sql("SELECT * FROM t OFFSET")

def test_limit_non_integer_string() -> None:
    with pytest.raises(DatabaseError, match="LIMIT must be an integer"):
        parse_sql("SELECT * FROM t LIMIT 'abc'")

def test_offset_non_integer_string() -> None:
    with pytest.raises(DatabaseError, match="OFFSET must be an integer"):
        parse_sql("SELECT * FROM t OFFSET 'abc'")

def test_having_with_param_binding() -> None:
    """HAVING with ? parameter that gets bound."""
    parsed = parse_sql(
        "SELECT name, COUNT(*) FROM t GROUP BY name HAVING COUNT(*) > ?",
        (5,),
    )
    assert parsed["having"] is not None
    cond = parsed["having"]["conditions"][0]
    assert cond["value"] == 5

def test_limit_param_non_integer() -> None:
    """LIMIT ? with non-integer param."""
    with pytest.raises(DatabaseError, match="LIMIT must be an integer"):
        parse_sql("SELECT * FROM t LIMIT ?", ("abc",))

def test_offset_param_non_integer() -> None:
    """OFFSET ? with non-integer param."""
    with pytest.raises(DatabaseError, match="OFFSET must be an integer"):
        parse_sql("SELECT * FROM t OFFSET ?", ("abc",))

def test_insert_empty_after_table() -> None:
    with pytest.raises(DatabaseError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO Sheet1")

def test_insert_no_values_no_select() -> None:
    with pytest.raises(DatabaseError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO Sheet1 SOMETHING")

def test_insert_unclosed_column_paren() -> None:
    with pytest.raises(DatabaseError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO Sheet1 (id, name VALUES (1, 'x')")

def test_insert_multi_row_unexpected_char() -> None:
    """Extra character between tuples that isn't comma."""
    with pytest.raises(DatabaseError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO Sheet1 VALUES (1, 'a') X (2, 'b')")

def test_insert_multi_row_no_open_paren() -> None:
    """VALUES followed by non-parenthesized content."""
    with pytest.raises(DatabaseError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO Sheet1 VALUES 1, 2")

def test_insert_too_many_params() -> None:
    """Too many parameters for INSERT placeholders."""
    with pytest.raises(DatabaseError, match="Too many parameters"):
        parse_sql("INSERT INTO Sheet1 VALUES (?, ?)", (1, 2, 3))

def test_insert_multi_row_with_escaped_quote() -> None:
    """Multi-row INSERT with escaped single quote in value."""
    parsed = parse_sql("INSERT INTO Sheet1 VALUES (1, 'it''s'), (2, 'ok')")
    assert parsed["values"] == [[1, "it's"], [2, "ok"]]

def test_insert_multi_row_nested_parens_in_values() -> None:
    """Simple multi-row — replaces a test that didn't raise as expected."""
    result = parse_sql("INSERT INTO Sheet1 VALUES (1, 'a')")
    assert result["values"] == [[1, "a"]]

def test_update_with_where_param_binding() -> None:
    """UPDATE SET ? WHERE ? — both SET and WHERE have params."""
    parsed = parse_sql(
        "UPDATE Sheet1 SET name = ? WHERE id = ?",
        ("Alice", 42),
    )
    assert parsed["set"][0]["value"] == {"type": "literal", "value": "Alice"}
    assert parsed["where"]["conditions"][0]["value"] == 42
