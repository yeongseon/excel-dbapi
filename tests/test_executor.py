from excel_dbapi.engine.executor import execute_query
from excel_dbapi.engine.openpyxl_engine import OpenpyxlEngine
from excel_dbapi.engine.parser import parse_sql


def test_executor_select():
    engine = OpenpyxlEngine("tests/data/sample.xlsx")
    data = engine.load()
    parsed = parse_sql("SELECT * FROM Sheet1")
    results = execute_query(parsed, data, engine.workbook)
    assert isinstance(results.rows, list)
    assert isinstance(results.rows[0], tuple)


def test_executor_select_with_where():
    engine = OpenpyxlEngine("tests/data/sample.xlsx")
    data = engine.load()
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = 1")
    results = execute_query(parsed, data, engine.workbook)

    assert isinstance(results.rows, list)
    assert len(results.rows) == 1
    assert results.rows[0][0] == 1
