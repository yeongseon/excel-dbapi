import pytest

from excel_dbapi.engine.parser import parse_sql


@pytest.mark.parametrize(
    "query",
    [
        "SELECT * FROM Sheet1 JOIN Sheet2 ON Sheet1.id = Sheet2.id",
        "SELECT name, COUNT(*) FROM Sheet1 GROUP BY name",
        "SELECT name, COUNT(*) FROM Sheet1 GROUP BY name HAVING COUNT(*) > 1",
        "SELECT DISTINCT name FROM Sheet1",
        "SELECT * FROM Sheet1 WHERE id = (SELECT id FROM Sheet2)",
    ],
)
def test_unsupported_sql_grammar_is_rejected(query):
    with pytest.raises(ValueError):
        parse_sql(query)
