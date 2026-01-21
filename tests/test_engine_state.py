from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from excel_dbapi.engine.openpyxl_engine import OpenpyxlEngine
from excel_dbapi.engine.pandas_engine import PandasEngine


def _create_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    wb.save(path)


def test_openpyxl_snapshot_restore(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    engine = OpenpyxlEngine(str(file_path))
    snapshot = engine.snapshot()
    engine.workbook = None
    engine.restore(snapshot)
    assert "Sheet1" in engine.data


def test_pandas_snapshot_restore(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    df = pd.DataFrame([{"id": 1, "name": "Alice"}])
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    engine = PandasEngine(str(file_path))
    snapshot = engine.snapshot()
    engine.data["Sheet1"].loc[0, "name"] = "Bob"
    engine.restore(snapshot)
    assert engine.data["Sheet1"].loc[0, "name"] == "Alice"


def test_openpyxl_save_without_workbook(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)
    engine = OpenpyxlEngine(str(file_path))
    engine.workbook = None
    with pytest.raises(ValueError):
        engine.save()
