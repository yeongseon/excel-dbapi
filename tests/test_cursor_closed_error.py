"""Tests for cursor.check_closed InterfaceError (issue #94).

PEP 249 requires InterfaceError for operations on closed cursors
and closed connections, not ProgrammingError.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import InterfaceError


def _make_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "t"
    ws.append(["id"])
    ws.append([1])
    wb.save(path)


def test_closed_cursor_raises_interface_error(tmp_path: Path) -> None:
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path)) as conn:
        cursor = conn.cursor()
        cursor.close()
        with pytest.raises(InterfaceError, match="Cursor is already closed"):
            cursor.execute("SELECT * FROM t")


def test_closed_connection_raises_interface_error(tmp_path: Path) -> None:
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    conn = ExcelConnection(str(file_path))
    cursor = conn.cursor()
    conn.close()
    with pytest.raises(InterfaceError, match="Cannot operate on a closed connection"):
        cursor.execute("SELECT * FROM t")


def test_closed_connection_fetchone_raises_interface_error(tmp_path: Path) -> None:
    file_path = tmp_path / "test.xlsx"
    _make_workbook(file_path)

    conn = ExcelConnection(str(file_path))
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM t")
    conn.close()
    with pytest.raises(InterfaceError, match="Cannot operate on a closed connection"):
        cursor.fetchone()
