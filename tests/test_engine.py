import os
import tempfile

import pandas as pd
import pytest

from excel_dbapi.engine.openpyxl_engine import OpenpyxlEngine
from excel_dbapi.engine.pandas_engine import PandasEngine


@pytest.fixture
def sample_excel_file():
    data = {
        "Sheet1": pd.DataFrame([{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
    }
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = tmp.name
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sheet_name, df in data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    yield path
    os.remove(path)


def test_pandas_engine_load_and_execute(sample_excel_file):
    engine = PandasEngine(sample_excel_file)
    assert "Sheet1" in engine.data
    result = engine.execute("SELECT * FROM [Sheet1$]")
    assert len(result) == 2
    assert result[0]["id"] == 1
    assert result[1]["name"] == "Bob"


def test_pandas_engine_save(sample_excel_file):
    engine = PandasEngine(sample_excel_file)
    # Modify data and save
    engine.data["Sheet1"].loc[0, "name"] = "Charlie"
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        save_path = tmp.name
    engine.save(save_path)

    # Reload and verify
    new_engine = PandasEngine(save_path)
    result = new_engine.execute("SELECT * FROM [Sheet1$]")
    assert result[0]["name"] == "Charlie"
    os.remove(save_path)


def test_openpyxl_engine_load_and_execute(sample_excel_file):
    engine = OpenpyxlEngine(sample_excel_file)
    assert "Sheet1" in engine.data
    result = engine.execute("SELECT * FROM [Sheet1$]")
    assert len(result) == 2
    assert result[0]["id"] == 1
    assert result[1]["name"] == "Bob"
