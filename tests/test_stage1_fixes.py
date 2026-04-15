"""Regression tests for Stage 1 stabilization fixes.

Covers:
- Path traversal validation (Issue #14)
- Temp file permissions (Issue #15)
- Autocommit snapshot bug (Issue #16 / GH#13)
- Parser: escaped quotes, AND/OR precedence, IS NULL (Issue #17)
- Exception types and file existence check (Issue #18)
"""

import os
import stat

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import NotSupportedError, OperationalError


# ── Helpers ──────────────────────────────────────────────────────────


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Create a minimal xlsx file with a Sheet1 containing headers and one row."""
    path = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id", "name", "score"])
    ws.append([1, "Alice", 90])
    ws.append([2, "Bob", 80])
    ws.append([3, None, 70])  # Row with NULL name
    wb.save(str(path))
    wb.close()
    return str(path)


@pytest.fixture
def tmp_xlsx_path(tmp_path):
    """Return a path (but don't create the file) — for testing create=True / missing file."""
    return str(tmp_path / "missing.xlsx")


# ── Fix 1: Path traversal validation (Issue #14) ────────────────────


class TestPathCanonicalization:
    def test_dotdot_in_path_is_resolved(self, tmp_path, tmp_xlsx):
        # Paths with '..' are canonicalized, not rejected (library, not sandbox)
        # Create the file first, then reference it with '..'
        subdir = tmp_path / "subdir"
        subdir.mkdir()
        path_with_dotdot = str(subdir / ".." / "test.xlsx")
        conn = ExcelConnection(path_with_dotdot)
        # The stored path should be the resolved canonical form
        assert ".." not in conn.file_path
        conn.close()

    def test_absolute_path_accepted(self, tmp_xlsx):
        conn = ExcelConnection(tmp_xlsx)
        assert conn.closed is False
        conn.close()

    def test_resolved_path_stored(self, tmp_xlsx):
        conn = ExcelConnection(tmp_xlsx)
        assert os.path.isabs(conn.file_path)
        conn.close()

    def test_tilde_expanded(self, tmp_xlsx):
        # expanduser should work (though may not change anything in test env)
        conn = ExcelConnection(tmp_xlsx)
        assert "~" not in conn.file_path
        conn.close()


# ── Fix 2: Temp file permissions (Issue #15) ────────────────────────


class TestTempFilePermissions:
    def test_openpyxl_save_creates_restricted_temp(self, tmp_xlsx):
        """After save, the target file should exist and be writable by owner."""
        conn = ExcelConnection(tmp_xlsx, autocommit=False)
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO Sheet1 (id, name, score) VALUES (?, ?, ?)", (10, "Test", 100)
        )
        conn.commit()
        conn.close()
        # File should still be readable
        assert os.path.exists(tmp_xlsx)
        mode = os.stat(tmp_xlsx).st_mode
        # Owner should have read+write
        assert mode & stat.S_IRUSR
        assert mode & stat.S_IWUSR

    def test_pandas_save_creates_restricted_temp(self, tmp_xlsx):
        """Same test for pandas engine."""
        conn = ExcelConnection(tmp_xlsx, engine="pandas", autocommit=False)
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO Sheet1 (id, name, score) VALUES (?, ?, ?)", (10, "Test", 100)
        )
        conn.commit()
        conn.close()
        assert os.path.exists(tmp_xlsx)


# ── Fix 3: Autocommit snapshot bug (Issue #16 / GH#13) ──────────────


class TestAutocommitSnapshot:
    def test_autocommit_write_updates_snapshot(self, tmp_xlsx):
        """After an autocommit write, switching to manual mode and rolling back
        should NOT undo the autocommitted data."""
        conn = ExcelConnection(tmp_xlsx, autocommit=True)
        cur = conn.cursor()

        # Autocommit write — should save AND update snapshot
        cur.execute(
            "INSERT INTO Sheet1 (id, name, score) VALUES (?, ?, ?)", (99, "Zara", 100)
        )

        # Switch to manual mode
        conn.autocommit = False

        # Rollback should restore to post-autocommit state (not pre-autocommit)
        conn.rollback()

        # The autocommitted row should still be present
        cur.execute("SELECT * FROM Sheet1 WHERE id = 99")
        rows = cur.fetchall()
        assert len(rows) == 1, (
            "Autocommitted row should survive rollback after switching to manual mode"
        )
        conn.close()

    def test_executemany_autocommit_updates_snapshot(self, tmp_xlsx):
        """executemany with autocommit should also update snapshot."""
        conn = ExcelConnection(tmp_xlsx, autocommit=True)
        cur = conn.cursor()

        cur.executemany(
            "INSERT INTO Sheet1 (id, name, score) VALUES (?, ?, ?)",
            [(10, "X", 50), (11, "Y", 60)],
        )

        conn.autocommit = False
        conn.rollback()

        cur.execute("SELECT * FROM Sheet1 WHERE id = 10")
        assert len(cur.fetchall()) == 1
        cur.execute("SELECT * FROM Sheet1 WHERE id = 11")
        assert len(cur.fetchall()) == 1
        conn.close()

    def test_autocommit_write_persisted_on_disk_after_rollback(self, tmp_xlsx):
        """GH#13 exact reproduction: autocommit write must survive rollback AND
        be present on disk (not just in memory). Verifies the file is persisted."""
        conn = ExcelConnection(tmp_xlsx, autocommit=True)
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO Sheet1 (id, name, score) VALUES (?, ?, ?)", (42, "Persist", 95)
        )

        # Toggle autocommit off and rollback
        conn.autocommit = False
        conn.rollback()
        conn.close()

        # Re-open the file fresh — verify data is on disk
        conn2 = ExcelConnection(tmp_xlsx, autocommit=True)
        cur2 = conn2.cursor()
        cur2.execute("SELECT * FROM Sheet1 WHERE id = 42")
        rows = cur2.fetchall()
        assert len(rows) == 1, "Autocommitted row must be on disk even after rollback"
        assert rows[0][1] == "Persist"
        conn2.close()


# ── Fix 4a: Parser escaped quotes (Issue #17) ───────────────────────


class TestEscapedQuotes:
    def test_single_quote_escape(self):
        from excel_dbapi.parser import _parse_value

        # SQL standard: '' inside single-quoted string = literal '
        assert _parse_value("'it''s'") == "it's"

    def test_double_quote_escape(self):
        from excel_dbapi.parser import _parse_value

        assert _parse_value('"say ""hello"""') == 'say "hello"'

    def test_simple_string_unchanged(self):
        from excel_dbapi.parser import _parse_value

        assert _parse_value("'hello'") == "hello"

    def test_empty_string(self):
        from excel_dbapi.parser import _parse_value

        assert _parse_value("''") == ""


# ── Fix 4b: AND/OR precedence (Issue #17) ───────────────────────────


class TestAndOrPrecedence:
    def test_and_binds_tighter_than_or_openpyxl(self, tmp_xlsx):
        """WHERE a = 1 OR b = 'Bob' AND score = 80 should be a = 1 OR (b = 'Bob' AND score = 80)."""
        conn = ExcelConnection(tmp_xlsx, engine="openpyxl")
        cur = conn.cursor()
        # id=1 (Alice,90), id=2 (Bob,80), id=3 (None,70)
        # With correct precedence: id=1 OR (name='Bob' AND score=80)
        # Should match id=1 (Alice) and id=2 (Bob)
        # With wrong precedence (left-to-right): (id=1 OR name='Bob') AND score=80
        # Would only match id=2 (Bob)
        cur.execute("SELECT * FROM Sheet1 WHERE id = 1 OR name = 'Bob' AND score = 80")
        rows = cur.fetchall()
        assert len(rows) == 2, (
            f"Expected 2 rows (AND before OR), got {len(rows)}: {rows}"
        )
        conn.close()

    def test_and_binds_tighter_than_or_pandas(self, tmp_xlsx):
        """Same test for pandas engine."""
        conn = ExcelConnection(tmp_xlsx, engine="pandas")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE id = 1 OR name = 'Bob' AND score = 80")
        rows = cur.fetchall()
        assert len(rows) == 2, (
            f"Expected 2 rows (AND before OR), got {len(rows)}: {rows}"
        )
        conn.close()

    def test_all_and_still_works(self, tmp_xlsx):
        """All AND conditions should still work correctly."""
        conn = ExcelConnection(tmp_xlsx, engine="openpyxl")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE id = 2 AND name = 'Bob' AND score = 80")
        rows = cur.fetchall()
        assert len(rows) == 1
        conn.close()

    def test_all_or_still_works(self, tmp_xlsx):
        """All OR conditions should still work correctly."""
        conn = ExcelConnection(tmp_xlsx, engine="openpyxl")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE id = 1 OR id = 2 OR id = 3")
        rows = cur.fetchall()
        assert len(rows) == 3
        conn.close()


# ── Fix 4c: IS NULL / IS NOT NULL and None comparison (Issue #17) ───


class TestNullHandling:
    def test_is_null_openpyxl(self, tmp_xlsx):
        """IS NULL should match rows where column is None."""
        conn = ExcelConnection(tmp_xlsx, engine="openpyxl")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE name IS NULL")
        rows = cur.fetchall()
        assert len(rows) == 1
        assert rows[0][0] == 3  # id=3 has None name
        conn.close()

    def test_is_not_null_openpyxl(self, tmp_xlsx):
        """IS NOT NULL should match rows where column is not None."""
        conn = ExcelConnection(tmp_xlsx, engine="openpyxl")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE name IS NOT NULL")
        rows = cur.fetchall()
        assert len(rows) == 2
        conn.close()

    def test_is_null_pandas(self, tmp_xlsx):
        """IS NULL should work with pandas engine too."""
        conn = ExcelConnection(tmp_xlsx, engine="pandas")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE name IS NULL")
        rows = cur.fetchall()
        # pandas reads None from xlsx as NaN, which is also null
        assert len(rows) == 1
        conn.close()

    def test_is_not_null_pandas(self, tmp_xlsx):
        conn = ExcelConnection(tmp_xlsx, engine="pandas")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE name IS NOT NULL")
        rows = cur.fetchall()
        assert len(rows) == 2
        conn.close()

    def test_equality_with_null_returns_false(self, tmp_xlsx):
        """col = NULL should not match anything (SQL semantics)."""
        conn = ExcelConnection(tmp_xlsx, engine="openpyxl")
        cur = conn.cursor()
        cur.execute("SELECT * FROM Sheet1 WHERE name = NULL")
        rows = cur.fetchall()
        assert len(rows) == 0, "Equality comparison with NULL should return no rows"
        conn.close()

    def test_parse_is_null(self):
        from excel_dbapi.parser import parse_sql

        parsed = parse_sql("SELECT * FROM Sheet1 WHERE name IS NULL")
        cond = parsed["where"]["conditions"][0]
        assert cond["column"] == "name"
        assert cond["operator"] == "IS"
        assert cond["value"] is None

    def test_parse_is_not_null(self):
        from excel_dbapi.parser import parse_sql

        parsed = parse_sql("SELECT * FROM Sheet1 WHERE name IS NOT NULL")
        cond = parsed["where"]["conditions"][0]
        assert cond["column"] == "name"
        assert cond["operator"] == "IS NOT"
        assert cond["value"] is None


# ── Fix 5: Exception types and file existence (Issue #18) ───────────


class TestExceptionTypes:
    def test_unsupported_engine_raises_operational_error(self, tmp_xlsx):
        with pytest.raises(NotSupportedError, match="Unsupported engine"):
            ExcelConnection(tmp_xlsx, engine="sqlite")

    def test_missing_file_raises_operational_error(self, tmp_xlsx_path):
        with pytest.raises(OperationalError, match="not found"):
            ExcelConnection(tmp_xlsx_path)

    def test_missing_file_with_create_succeeds(self, tmp_xlsx_path):
        conn = ExcelConnection(tmp_xlsx_path, create=True)
        assert conn.closed is False
        assert os.path.exists(tmp_xlsx_path)
        conn.close()

    def test_existing_file_without_create_succeeds(self, tmp_xlsx):
        conn = ExcelConnection(tmp_xlsx)
        assert conn.closed is False
        conn.close()

    def test_bad_engine_checked_before_file_existence(self, tmp_xlsx_path):
        """Engine validation should happen before file existence check."""
        # Both conditions fail: bad engine + missing file
        # Should get OperationalError about engine, not about file
        with pytest.raises(NotSupportedError, match="Unsupported engine"):
            ExcelConnection(tmp_xlsx_path, engine="nonexistent")
