from excel_dbapi.engines.openpyxl.backend import OpenpyxlBackend
from excel_dbapi.executor import SharedExecutor
from excel_dbapi.parser import parse_sql


def test_executor_select():
    engine = OpenpyxlBackend("tests/data/sample.xlsx")
    parsed = parse_sql("SELECT * FROM Sheet1")
    results = SharedExecutor(engine).execute(parsed)
    assert isinstance(results.rows, list)
    assert isinstance(results.rows[0], tuple)


def test_executor_select_with_where():
    engine = OpenpyxlBackend("tests/data/sample.xlsx")
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = 1")
    results = SharedExecutor(engine).execute(parsed)

    assert isinstance(results.rows, list)
    assert len(results.rows) == 1
    assert results.rows[0][0] == 1
