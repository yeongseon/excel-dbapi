import pytest

from excel_dbapi.engine.parser import parse_sql


def test_parse_select_query():
    query = "SELECT * FROM [Sheet1$] WHERE id = ?"
    params = (1,)
    result = parse_sql(query, params)
    assert result["action"] == "SELECT"
    assert result["table"] == "sheet1"
    assert result["where"] == "id == 1"


def test_parse_insert_query():
    query = "INSERT INTO [Sheet1$] VALUES (?, ?)"
    params = (1, "Alice")
    result = parse_sql(query, params)
    assert result["action"] == "INSERT"
    assert result["table"] == "sheet1"


def test_parse_query_invalid_params():
    query = "SELECT * FROM [Sheet1$] WHERE id = ?"
    params = (1, 2)
    with pytest.raises(ValueError):
        parse_sql(query, params)


def test_parse_query_not_supported():
    query = "DELETE FROM [Sheet1$]"
    with pytest.raises(NotImplementedError):
        parse_sql(query)


import pytest

from excel_dbapi.engine.parser import parse_sql


def test_parse_select_no_where():
    query = "SELECT * FROM [Sheet1$]"
    result = parse_sql(query)
    assert result["action"] == "SELECT"
    assert result["table"] == "sheet1"
    assert result["where"] is None


def test_parse_insert_invalid():
    query = "INSERT INTO [Sheet1$]"
    result = parse_sql(query)
    assert result["action"] == "INSERT"
    assert result["table"] == "sheet1"
