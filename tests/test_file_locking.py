import os
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import OperationalError


def _create_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    wb.save(path)


def test_autocommit_false_holds_write_lock(tmp_path: Path) -> None:
    file_path = tmp_path / "lock.xlsx"
    _create_workbook(file_path)

    conn1 = ExcelConnection(str(file_path), engine="openpyxl", autocommit=False)
    try:
        with pytest.raises(OperationalError, match="File is locked by another process"):
            ExcelConnection(str(file_path), engine="openpyxl", autocommit=False)
    finally:
        conn1.close()


def test_autocommit_true_acquires_lock_on_first_write(tmp_path: Path) -> None:
    file_path = tmp_path / "lock.xlsx"
    _create_workbook(file_path)

    conn1 = ExcelConnection(str(file_path), engine="openpyxl", autocommit=True)
    conn2 = ExcelConnection(str(file_path), engine="openpyxl", autocommit=False)
    conn2.close()

    try:
        conn1.execute("INSERT INTO Sheet1 (id, name) VALUES (2, 'Bob')")
        with pytest.raises(OperationalError, match="File is locked by another process"):
            ExcelConnection(str(file_path), engine="openpyxl", autocommit=False)
    finally:
        conn1.close()


def test_file_locking_can_be_disabled(tmp_path: Path) -> None:
    file_path = tmp_path / "lock.xlsx"
    _create_workbook(file_path)

    conn1 = ExcelConnection(
        str(file_path), engine="openpyxl", autocommit=False, file_locking=False
    )
    conn2 = ExcelConnection(
        str(file_path), engine="openpyxl", autocommit=False, file_locking=False
    )

    conn1.close()
    conn2.close()


def test_stale_lock_file_is_recovered(tmp_path: Path) -> None:
    file_path = tmp_path / "stale-lock.xlsx"
    _create_workbook(file_path)
    lock_path = Path(f"{file_path}.lock")
    lock_path.write_text("999999", encoding="ascii")

    conn = ExcelConnection(str(file_path), engine="openpyxl", autocommit=False)
    try:
        assert conn.closed is False
        lock_pid = int(lock_path.read_text(encoding="ascii"))
        assert lock_pid == os.getpid()
    finally:
        conn.close()


def test_corrupted_binary_lock_file_recovery(tmp_path: Path) -> None:
    """Lock file with non-ASCII bytes should not crash _clear_stale_lock."""
    file_path = tmp_path / "data.xlsx"
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id"])
    ws.append([1])
    wb.save(file_path)

    lock_path = Path(f"{file_path}.lock")
    # Write binary garbage that would cause UnicodeDecodeError with ascii encoding
    lock_path.write_bytes(b"\xff\xfe\x00\x80corrupted")

    # Should not raise — lock is treated as stale and replaced
    conn = ExcelConnection(str(file_path), engine="openpyxl", autocommit=False)
    try:
        assert conn.closed is False
    finally:
        conn.close()
