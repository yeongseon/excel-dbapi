from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection


def _create_typed_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id", "event_date", "event_ts", "is_active"])
    ws.append([1, date(2024, 1, 2), datetime(2024, 1, 2, 8, 30, 0), True])
    ws.append([2, date(2023, 12, 31), datetime(2023, 12, 31, 23, 59, 0), False])
    ws.append([3, None, None, None])
    ws.append([4, date(2024, 1, 1), datetime(2024, 1, 1, 12, 0, 0), True])
    ws.append([5, date(2024, 1, 3), datetime(2024, 1, 3, 0, 0, 0), False])
    wb.save(path)


def _query(file_path: Path, query: str, params: tuple[object, ...] | None = None) -> list[tuple[object, ...]]:
    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        return cursor.fetchall()


def test_order_by_date_and_datetime_native(tmp_path: Path) -> None:
    file_path = tmp_path / "typed_values.xlsx"
    _create_typed_workbook(file_path)

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE event_date IS NOT NULL ORDER BY event_date ASC",
    )
    assert rows == [(2,), (4,), (1,), (5,)]

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE event_ts IS NOT NULL ORDER BY event_ts DESC",
    )
    assert rows == [(5,), (1,), (4,), (2,)]


def test_where_date_predicates_with_string_literals(tmp_path: Path) -> None:
    file_path = tmp_path / "typed_values.xlsx"
    _create_typed_workbook(file_path)

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE event_date > '2024-01-01' ORDER BY id",
    )
    assert rows == [(1,), (5,)]

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE event_date BETWEEN '2024-01-01' AND '2024-01-02' ORDER BY id",
    )
    assert rows == [(1,), (4,)]


def test_boolean_ordering_and_where_comparison(tmp_path: Path) -> None:
    file_path = tmp_path / "typed_values.xlsx"
    _create_typed_workbook(file_path)

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE is_active IS NOT NULL ORDER BY is_active ASC, id ASC",
    )
    assert rows == [(2,), (5,), (1,), (4,)]

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE is_active > ? ORDER BY id",
        (False,),
    )
    assert rows == [(1,), (4,)]


def test_null_handling_with_date_and_boolean_ordering(tmp_path: Path) -> None:
    file_path = tmp_path / "typed_values.xlsx"
    _create_typed_workbook(file_path)

    rows = _query(file_path, "SELECT id FROM Sheet1 ORDER BY event_date ASC")
    assert rows == [(2,), (4,), (1,), (5,), (3,)]

    rows = _query(file_path, "SELECT id FROM Sheet1 ORDER BY is_active ASC, id ASC")
    assert rows == [(2,), (5,), (1,), (4,), (3,)]


def test_mixed_date_datetime_comparisons(tmp_path: Path) -> None:
    file_path = tmp_path / "typed_values.xlsx"
    _create_typed_workbook(file_path)

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE event_date <= '2024-01-02 00:00:00' AND event_date IS NOT NULL ORDER BY id",
    )
    assert rows == [(1,), (2,), (4,)]

    rows = _query(
        file_path,
        "SELECT id FROM Sheet1 WHERE event_ts >= '2024-01-02' ORDER BY id",
    )
    assert rows == [(1,), (5,)]
