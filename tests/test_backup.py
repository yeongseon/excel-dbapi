"""Tests for backup and warn_rows connection options."""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl
import pytest

from excel_dbapi import connect


def _make_workbook(path: Path, rows: int = 3) -> Path:
    """Create a minimal .xlsx workbook with the given number of data rows."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    for i in range(1, rows + 1):
        ws.append([i, f"row{i}"])
    wb.save(str(path))
    return path


# ── Backup tests ────────────────────────────────────────────────────


class TestBackupCreation:
    """Backup feature creates a timestamped copy before the first mutating op."""

    def test_backup_created_on_first_mutation(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx")
        backup_dir = tmp_path / ".excel-dbapi-backups"

        with connect(str(wb_path), backup=True) as conn:
            # No backup yet — no mutation has occurred
            assert not backup_dir.exists()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (99, 'new')")

        # Backup should now exist
        assert backup_dir.exists()
        backups = list(backup_dir.iterdir())
        assert len(backups) == 1
        assert backups[0].suffix == ".xlsx"
        assert backups[0].stem.startswith("test.")

    def test_backup_created_only_once_per_session(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx")
        backup_dir = tmp_path / ".excel-dbapi-backups"

        with connect(str(wb_path), backup=True) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (10, 'a')")
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (11, 'b')")
            cursor.execute("DELETE FROM Sheet1 WHERE id = 10")

        backups = list(backup_dir.iterdir())
        assert len(backups) == 1

    def test_no_backup_when_disabled(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx")
        backup_dir = tmp_path / ".excel-dbapi-backups"

        with connect(str(wb_path), backup=False) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (99, 'new')")

        assert not backup_dir.exists()

    def test_custom_backup_dir(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx")
        custom_dir = tmp_path / "my_backups"

        with connect(str(wb_path), backup=True, backup_dir=str(custom_dir)) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (99, 'new')")

        assert custom_dir.exists()
        backups = list(custom_dir.iterdir())
        assert len(backups) == 1
        assert backups[0].suffix == ".xlsx"

    def test_backup_filename_contains_timestamp(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx")

        with connect(str(wb_path), backup=True) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (99, 'new')")

        backup_dir = tmp_path / ".excel-dbapi-backups"
        backups = list(backup_dir.iterdir())
        assert len(backups) == 1
        # Pattern: test.YYYYMMDD-HHMMSS-ffffff.xlsx
        assert re.match(r"test\.\d{8}-\d{6}-\d{6}\.xlsx", backups[0].name)

    def test_backup_preserves_original_content(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx")
        original_bytes = wb_path.read_bytes()

        with connect(str(wb_path), backup=True) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (99, 'new')")

        backup_dir = tmp_path / ".excel-dbapi-backups"
        backups = list(backup_dir.iterdir())
        assert backups[0].read_bytes() == original_bytes


# ── warn_rows tests ─────────────────────────────────────────────────


class TestWarnRows:
    """warn_rows emits a UserWarning when a sheet exceeds the threshold."""

    def test_warning_emitted_when_rows_exceed_threshold(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx", rows=10)

        with pytest.warns(UserWarning, match=r"Sheet 'Sheet1' has \d+ rows"):
            with connect(str(wb_path), warn_rows=5) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM Sheet1")

    def test_no_warning_when_below_threshold(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx", rows=3)

        import warnings

        with warnings.catch_warnings():
            warnings.simplefilter("error", UserWarning)
            with connect(str(wb_path), warn_rows=100) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM Sheet1")

    def test_no_warning_when_warn_rows_is_none(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx", rows=10)

        import warnings

        with warnings.catch_warnings():
            warnings.simplefilter("error", UserWarning)
            with connect(str(wb_path), warn_rows=None) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM Sheet1")

    def test_warning_emitted_once_per_sheet(self, tmp_path: Path) -> None:
        wb_path = _make_workbook(tmp_path / "test.xlsx", rows=10)

        import warnings

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            with connect(str(wb_path), warn_rows=5) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM Sheet1")
                cursor.execute("SELECT * FROM Sheet1")

        warn_row_warnings = [
            x
            for x in w
            if issubclass(x.category, UserWarning) and "rows" in str(x.message)
        ]
        assert len(warn_row_warnings) == 1
