import pytest

from excel_dbapi.parser import parse_sql


@pytest.mark.parametrize(
    "query",
    [
        "SELECT * FROM Sheet1 JOIN Sheet2 ON Sheet1.id = Sheet2.id",
        "SELECT * FROM Sheet1 WHERE id = (SELECT id FROM Sheet2)",
    ],
)
def test_unsupported_sql_grammar_is_rejected(query):
    with pytest.raises(ValueError):
        parse_sql(query)


def test_rejects_union():
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users UNION SELECT * FROM admins")


def test_rejects_for_update():
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users FOR UPDATE")


def test_rejects_nulls_last():
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users ORDER BY id ASC NULLS LAST")


def test_rejects_intersect():
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users INTERSECT SELECT * FROM admins")


def test_rejects_except():
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        parse_sql("SELECT * FROM users EXCEPT SELECT * FROM admins")
