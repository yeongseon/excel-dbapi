from pathlib import Path
from typing import Any, cast

import httpx
import pytest
from openpyxl import Workbook

from excel_dbapi import connect
from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import Error, OperationalError, ProgrammingError


def _create_workbook(
    path: Path, sheet: str, headers: list[object], rows: list[list[object]]
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = sheet
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)
    workbook.close()


def test_graph_invalid_credential_is_wrapped_as_operational_error() -> None:
    with pytest.raises(OperationalError, match="Cannot normalise"):
        connect(
            "msgraph://drives/d/items/i",
            engine="graph",
            credential=cast(Any, 42),
        )


def test_graph_token_provider_failure_is_translated_during_execute() -> None:
    class ExplodingTokenProvider:
        def get_token(self, *args: Any) -> Any:
            del args
            raise RuntimeError("token boom")

    transport = httpx.MockTransport(lambda request: httpx.Response(200, json={}))
    with ExcelConnection(
        "msgraph://drives/drv-test/items/itm-test",
        engine="graph",
        credential=ExplodingTokenProvider(),
        transport=transport,
    ) as conn:
        cursor = conn.cursor()
        with pytest.raises(
            Error, match="Failed to acquire authentication token: token boom"
        ):
            cursor.execute("SELECT * FROM Employees")


def test_executor_resolves_unicode_sheet_names_with_casefold(tmp_path: Path) -> None:
    file_path = tmp_path / "unicode-sheet-casefold.xlsx"
    _create_workbook(file_path, "Straße", ["id"], [[1]])

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM STRASSE")
        assert cursor.fetchall() == [(1,)]


def test_lastrowid_is_cleared_after_failed_execute(tmp_path: Path) -> None:
    file_path = tmp_path / "lastrowid-reset.xlsx"
    _create_workbook(file_path, "Sheet1", ["id", "name"], [[1, "Alice"]])

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Sheet1 VALUES (2, 'Bob')")
        assert cursor.lastrowid is not None

        with pytest.raises(ProgrammingError):
            cursor.execute("SELECT * FROM MissingSheet")

        assert cursor.lastrowid is None
