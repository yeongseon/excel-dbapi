"""Tests for formula injection defense (sanitize.py + executor integration)."""

import pytest

from excel_dbapi.sanitize import sanitize_cell_value, sanitize_row


# ---------------------------------------------------------------------------
# Unit tests for sanitize_cell_value
# ---------------------------------------------------------------------------


class TestSanitizeCellValue:
    """sanitize_cell_value escapes formula-triggering prefixes."""

    @pytest.mark.parametrize(
        "raw, expected",
        [
            ("=SUM(A1:A10)", "'=SUM(A1:A10)"),
            ("+cmd|'/C calc'!A0", "'+cmd|'/C calc'!A0"),
            ("-1+1", "'-1+1"),
            ("@SUM(A1:A10)", "'@SUM(A1:A10)"),
            ("\tcmd", "'\tcmd"),
            ("\rcmd", "'\rcmd"),
        ],
    )
    def test_dangerous_prefixes_escaped(self, raw: str, expected: str) -> None:
        assert sanitize_cell_value(raw) == expected

    @pytest.mark.parametrize(
        "safe",
        [
            "hello",
            "123",
            "",
            "Alice",
            "normal text",
            "no formula here",
        ],
    )
    def test_safe_strings_unchanged(self, safe: str) -> None:
        assert sanitize_cell_value(safe) == safe

    @pytest.mark.parametrize("value", [42, 3.14, None, True, False])
    def test_non_strings_unchanged(self, value) -> None:
        assert sanitize_cell_value(value) is value


class TestSanitizeRow:
    """sanitize_row applies sanitize_cell_value to every element."""

    def test_mixed_row(self) -> None:
        row = ["=SUM(A1)", "safe", 42, None, "+cmd"]
        result = sanitize_row(row)
        assert result == ["'=SUM(A1)", "safe", 42, None, "'+cmd"]

    def test_empty_row(self) -> None:
        assert sanitize_row([]) == []

    def test_all_safe(self) -> None:
        row = ["a", "b", 1, 2]
        assert sanitize_row(row) == ["a", "b", 1, 2]


# ---------------------------------------------------------------------------
# Integration tests: formula injection defense in INSERT / UPDATE paths
# ---------------------------------------------------------------------------


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Create a minimal xlsx with a 'Sheet1' containing headers [id, name]."""
    from openpyxl import Workbook

    path = str(tmp_path / "test.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    wb.save(path)
    return path


class TestOpenpyxlSanitizationIntegration:
    """Formula injection defense works end-to-end with OpenpyxlEngine."""

    def test_insert_sanitizes_by_default(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(tmp_xlsx) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                (2, "=HYPERLINK('http://evil.com')"),
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = '2'")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "'=HYPERLINK('http://evil.com')"

    def test_insert_no_sanitize_when_disabled(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(tmp_xlsx, sanitize_formulas=False) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                (3, "=SUM(A1:A10)"),
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = '3'")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "=SUM(A1:A10)"

    def test_update_sanitizes_by_default(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(tmp_xlsx) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE Sheet1 SET name = '+cmd' WHERE id = 1",
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = '1'")
            row = cursor.fetchone()
            assert row is not None
            # +cmd starts with '+', so it gets sanitized with a leading quote
            assert row[1] == "'+cmd"

    def test_update_no_sanitize_when_disabled(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(tmp_xlsx, sanitize_formulas=False) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE Sheet1 SET name = '+cmd' WHERE id = 1",
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = '1'")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "+cmd"


class TestPandasSanitizationIntegration:
    """Formula injection defense works end-to-end with PandasEngine."""

    def test_insert_sanitizes_by_default(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(tmp_xlsx, engine="pandas") as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                (2, "=HYPERLINK('http://evil.com')"),
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = 2")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "'=HYPERLINK('http://evil.com')"

    def test_insert_no_sanitize_when_disabled(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(
            tmp_xlsx, engine="pandas", sanitize_formulas=False
        ) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                (3, "=SUM(A1:A10)"),
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = 3")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "=SUM(A1:A10)"

    def test_update_sanitizes_by_default(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(tmp_xlsx, engine="pandas") as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE Sheet1 SET name = '@SUM(A1)' WHERE id = 1",
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = 1")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "'@SUM(A1)"

    def test_update_no_sanitize_when_disabled(self, tmp_xlsx: str) -> None:
        from excel_dbapi.connection import ExcelConnection

        with ExcelConnection(
            tmp_xlsx, engine="pandas", sanitize_formulas=False
        ) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "UPDATE Sheet1 SET name = '@SUM(A1)' WHERE id = 1",
            )
            cursor.execute("SELECT id, name FROM Sheet1 WHERE id = 1")
            row = cursor.fetchone()
            assert row is not None
            assert row[1] == "@SUM(A1)"
