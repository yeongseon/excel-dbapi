"""Tests for executemany exception mapping (issue #92).

Verifies that executemany maps non-DB-API exceptions to the proper
DB-API exception types, matching the mapping used by execute().
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _make_workbook(path: Path) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "items"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    wb.save(path)


def test_executemany_maps_value_error(tmp_path: Path) -> None:
    """executemany should map ValueError to ProgrammingError."""
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        # Column 'bad_col' does not exist → ValueError → ProgrammingError
        with pytest.raises(ProgrammingError):
            cursor.executemany(
                "INSERT INTO items (id, bad_col) VALUES (?, ?)",
                [(2, "Bob")],
            )


def test_executemany_resets_state_on_error(tmp_path: Path) -> None:
    """After an executemany error, cursor state should be reset."""
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.executemany(
                "INSERT INTO items (id, bad_col) VALUES (?, ?)",
                [(2, "Bob")],
            )
        assert cursor.rowcount == -1
        assert cursor.lastrowid is None
        assert cursor.description is None


def test_executemany_rollback_on_error(tmp_path: Path) -> None:
    """Transactional backend should rollback on executemany error."""
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        # Trigger an error during executemany — bad column name
        with pytest.raises(ProgrammingError):
            cursor.executemany(
                "INSERT INTO items (id, bad_col) VALUES (?, ?)",
                [(2, "Bob"), (3, "Carol")],
            )

        # After rollback, original data should be intact
        cursor.execute("SELECT COUNT(*) FROM items")
        assert cursor.fetchone() == (1,)


def test_executemany_success(tmp_path: Path) -> None:
    """executemany should work correctly on success."""
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.executemany(
            "INSERT INTO items (id, name) VALUES (?, ?)",
            [(2, "Bob"), (3, "Carol")],
        )
        assert cursor.rowcount == 2
        cursor.execute("SELECT * FROM items ORDER BY id")
        rows = cursor.fetchall()
        assert len(rows) == 3
        assert rows[1] == (2, "Bob")
        assert rows[2] == (3, "Carol")
