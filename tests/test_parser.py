import pytest
from excel_dbapi.engine.parser import parse_sql


def test_parse_valid_sql():
    parsed = parse_sql("SELECT * FROM Sheet1")
    assert parsed["action"] == "SELECT"
    assert parsed["table"] == "Sheet1"
    assert parsed["columns"] == "*"


def test_parse_invalid_sql():
    with pytest.raises(ValueError):
        parse_sql("INVALID SQL")


def test_parse_sql_with_where():
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = 1")
    assert parsed["where"]["column"] == "id"
    assert parsed["where"]["operator"] == "="
    assert parsed["where"]["value"] == "1"
