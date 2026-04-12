from excel_dbapi.engines.openpyxl.backend import OpenpyxlBackend


def test_engine_load():
    engine = OpenpyxlBackend("tests/data/sample.xlsx")
    assert isinstance(engine.data, dict)
    assert "Sheet1" in engine.data


def test_engine_execute_select():
    engine = OpenpyxlBackend("tests/data/sample.xlsx")
    results = engine.execute("SELECT * FROM Sheet1")
    assert isinstance(results.rows, list)
    assert isinstance(results.rows[0], tuple)
