from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import OperationalError


def _create_workbook(path: Path, rows: int) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id", "payload"])
    for index in range(rows):
        ws.append([index + 1, f"row-{index + 1}"])
    wb.save(path)


@pytest.mark.parametrize("engine", ["openpyxl", "pandas"])
def test_row_limit_raises_operational_error(tmp_path: Path, engine: str) -> None:
    file_path = tmp_path / "limit.xlsx"
    _create_workbook(file_path, rows=5)

    with ExcelConnection(str(file_path), engine=engine, max_rows=4) as conn:
        with pytest.raises(OperationalError, match="Sheet exceeds max_rows limit"):
            conn.execute("SELECT * FROM Sheet1")


@pytest.mark.parametrize("engine", ["openpyxl", "pandas"])
def test_row_limit_warns_at_eighty_percent(tmp_path: Path, engine: str) -> None:
    file_path = tmp_path / "warn.xlsx"
    _create_workbook(file_path, rows=8)

    with ExcelConnection(str(file_path), engine=engine, max_rows=10) as conn:
        with pytest.warns(UserWarning, match="reached 8/10 rows"):
            conn.execute("SELECT * FROM Sheet1")


def test_memory_limit_raises_for_pandas_backend(tmp_path: Path) -> None:
    file_path = tmp_path / "mem.xlsx"
    df = pd.DataFrame({"id": [1, 2, 3], "payload": ["x" * 5000] * 3})
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    with ExcelConnection(str(file_path), engine="pandas", max_memory_mb=0.001) as conn:
        with pytest.raises(OperationalError, match="Sheet exceeds max_memory_mb limit"):
            conn.execute("SELECT * FROM Sheet1")
