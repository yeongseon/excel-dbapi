from __future__ import annotations

import datetime
from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.reflection import _infer_type, get_columns


def _write_sheet(path: Path, headers: list[str], rows: list[list[object | None]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def test_default_sample_size_is_100(tmp_path: Path) -> None:
    file_path = tmp_path / "default-sample-size.xlsx"
    rows: list[list[object | None]] = [[value] for value in range(100)]
    rows.extend([[f"text-{value}"] for value in range(100, 200)])
    _write_sheet(file_path, ["mixed"], rows)

    with ExcelConnection(str(file_path), engine="openpyxl") as connection:
        columns = get_columns(connection, "Sheet1")

    assert columns[0]["type"] == "INTEGER"


def test_custom_sample_size_and_full_scan(tmp_path: Path) -> None:
    file_path = tmp_path / "custom-sample-size.xlsx"
    rows: list[list[object | None]] = [[value] for value in range(100)]
    rows.extend([[f"text-{value}"] for value in range(100, 200)])
    _write_sheet(file_path, ["mixed"], rows)

    with ExcelConnection(str(file_path), engine="openpyxl") as connection:
        columns_150 = get_columns(connection, "Sheet1", sample_size=150)
        columns_full = get_columns(connection, "Sheet1", sample_size=None)

    assert columns_150[0]["type"] == "TEXT"
    assert columns_full[0]["type"] == "TEXT"


def test_infer_type_prefers_float_for_int_float_mix() -> None:
    assert _infer_type([1, 2, 3.5, None])["type"] == "FLOAT"


def test_infer_type_prefers_datetime_for_date_datetime_mix() -> None:
    assert _infer_type(
        [
            datetime.date(2024, 1, 1),
            datetime.datetime(2024, 1, 2, 8, 30, 0),
            None,
        ]
    )["type"] == "DATETIME"


def test_infer_type_returns_text_for_truly_mixed_values() -> None:
    assert _infer_type([1, True, "text", None])["type"] == "TEXT"


def test_infer_type_empty_column_is_text() -> None:
    assert _infer_type([None, None])["type"] == "TEXT"


def test_sparse_data_keeps_dominant_non_null_type(tmp_path: Path) -> None:
    file_path = tmp_path / "sparse-column.xlsx"
    rows: list[list[object | None]] = [[None] for _ in range(90)]
    rows.extend([[value] for value in range(10)])
    _write_sheet(file_path, ["sparse"], rows)

    with ExcelConnection(str(file_path), engine="openpyxl") as connection:
        columns = get_columns(connection, "Sheet1", sample_size=None)

    assert columns[0]["type"] == "INTEGER"
    assert columns[0]["nullable"] is True
