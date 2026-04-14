from __future__ import annotations

import datetime
from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.reflection import (
    METADATA_SHEET,
    _infer_type,
    get_columns,
    has_table,
    list_tables,
    read_table_metadata,
    remove_table_metadata,
    write_table_metadata,
)


def _make_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "People"
    ws.append(
        [
            "id",
            "name",
            "active",
            "score",
            "created_date",
            "created_at",
            "notes",
        ]
    )
    ws.append(
        [
            1,
            "Alice",
            True,
            98.5,
            datetime.date(2024, 1, 2),
            datetime.datetime(2024, 1, 2, 10, 0, 0),
            None,
        ]
    )
    ws.append(
        [
            2,
            "Bob",
            False,
            77,
            datetime.date(2024, 1, 3),
            datetime.datetime(2024, 1, 3, 11, 30, 0),
            "ok",
        ]
    )
    wb.save(path)


def test_list_tables_and_has_table(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        tables = list_tables(conn)
        assert tables == ["People"]
        assert has_table(conn, "people") is True
        assert has_table(conn, "PEOPLE") is True
        assert has_table(conn, "missing") is False

        write_table_metadata(
            conn,
            "People",
            [
                {
                    "name": "id",
                    "type_name": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                }
            ],
        )
        tables_with_meta = list_tables(conn, include_meta=True)
        assert "People" in tables_with_meta
        assert METADATA_SHEET in tables_with_meta


def test_get_columns_with_mixed_types(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection-columns.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        columns = get_columns(conn, "People")

    by_name = {column["name"]: column for column in columns}
    assert by_name["id"]["type"] == "INTEGER"
    assert by_name["id"]["type_name"] == "INTEGER"
    assert by_name["name"]["type"] == "TEXT"
    assert by_name["name"]["type_name"] == "TEXT"
    assert by_name["active"]["type"] == "BOOLEAN"
    assert by_name["score"]["type"] == "FLOAT"
    assert by_name["created_date"]["type"] == "DATETIME"
    assert by_name["created_at"]["type"] == "DATETIME"
    assert by_name["notes"]["type"] == "TEXT"
    assert by_name["notes"]["nullable"] is True


def test_infer_type_combinations() -> None:
    assert _infer_type([None, None]) == {"type": "TEXT", "nullable": True}
    assert _infer_type([1, 2, 3]) == {"type": "INTEGER", "nullable": False}
    assert _infer_type([1, 2.5]) == {"type": "FLOAT", "nullable": False}
    assert _infer_type([True, False, None]) == {"type": "BOOLEAN", "nullable": True}
    assert _infer_type([datetime.date(2024, 1, 1)]) == {
        "type": "DATE",
        "nullable": False,
    }
    assert _infer_type([datetime.datetime(2024, 1, 1, 1, 0, 0)]) == {
        "type": "DATETIME",
        "nullable": False,
    }
    assert _infer_type(
        [datetime.date(2024, 1, 1), datetime.datetime(2024, 1, 1, 1, 0, 0)]
    ) == {
        "type": "DATETIME",
        "nullable": False,
    }
    assert _infer_type([1, "x"]) == {"type": "TEXT", "nullable": False}


def test_write_read_remove_table_metadata(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection-meta.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        assert read_table_metadata(conn, "People") is None
        assert read_table_metadata(conn, "Missing") is None

        write_table_metadata(
            conn,
            "People",
            [
                {
                    "name": "id",
                    "type_name": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                },
                {
                    "name": "name",
                    "type_name": "TEXT",
                    "nullable": True,
                    "primary_key": False,
                },
            ],
        )
        meta = read_table_metadata(conn, "People")
        assert meta == [
            {
                "name": "id",
                "type_name": "INTEGER",
                "nullable": False,
                "primary_key": True,
            },
            {
                "name": "name",
                "type_name": "TEXT",
                "nullable": True,
                "primary_key": False,
            },
        ]

        write_table_metadata(
            conn,
            "People",
            [
                {
                    "name": "id",
                    "type_name": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                }
            ],
        )
        updated_meta = read_table_metadata(conn, "People")
        assert updated_meta == [
            {
                "name": "id",
                "type_name": "INTEGER",
                "nullable": False,
                "primary_key": True,
            }
        ]

        remove_table_metadata(conn, "People")
        assert read_table_metadata(conn, "People") is None


def test_remove_table_metadata_missing_sheet(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection-no-meta.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        remove_table_metadata(conn, "People")
        assert list_tables(conn) == ["People"]


def test_get_columns_case_insensitive_table_resolution(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection-case-insensitive.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        columns = get_columns(conn, "people")

    assert [column["name"] for column in columns][:3] == ["id", "name", "active"]


def test_write_metadata_accepts_legacy_type_key(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection-type-key.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        write_table_metadata(
            conn,
            "People",
            [
                {
                    "name": "id",
                    "type": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                }
            ],
        )
        metadata = read_table_metadata(conn, "People")

    assert metadata == [
        {
            "name": "id",
            "type_name": "INTEGER",
            "nullable": False,
            "primary_key": True,
        }
    ]


def test_table_metadata_operations_are_case_insensitive(tmp_path: Path) -> None:
    file_path = tmp_path / "reflection-metadata-case-insensitive.xlsx"
    _make_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        write_table_metadata(
            conn,
            "People",
            [
                {
                    "name": "id",
                    "type_name": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                }
            ],
        )

        metadata = read_table_metadata(conn, "people")
        assert metadata == [
            {
                "name": "id",
                "type_name": "INTEGER",
                "nullable": False,
                "primary_key": True,
            }
        ]

        remove_table_metadata(conn, "PEOPLE")
        assert read_table_metadata(conn, "People") is None
