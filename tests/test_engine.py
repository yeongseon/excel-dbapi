import pytest
from excel_dbapi.engine.openpyxl_engine import OpenpyxlEngine


def test_engine_load():
    engine = OpenpyxlEngine("tests/data/sample.xlsx")
    data = engine.load()
    assert isinstance(data, dict)
    assert "Sheet1" in data


def test_engine_execute_select():
    engine = OpenpyxlEngine("tests/data/sample.xlsx")
    results = engine.execute("SELECT * FROM Sheet1")
    assert isinstance(results.rows, list)
    assert isinstance(results.rows[0], tuple)
