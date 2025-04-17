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
