from excel_dbapi.engine.executor import execute_query
from excel_dbapi.engine.openpyxl_engine import OpenpyxlEngine
from excel_dbapi.engine.parser import parse_sql


def test_executor_select():
    engine = OpenpyxlEngine("tests/data/sample.xlsx")
    data = engine.load()
    parsed = parse_sql("SELECT * FROM Sheet1")
    results = execute_query(parsed, data)
    assert isinstance(results, list)
    assert isinstance(results[0], dict)


def test_executor_select_with_where():
    engine = OpenpyxlEngine("tests/data/sample.xlsx")
    data = engine.load()
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = 1")
    results = execute_query(parsed, data)

    assert isinstance(results, list)
    assert len(results) == 1
    assert results[0]["id"] == 1