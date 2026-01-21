import pytest
from excel_dbapi.engine.parser import parse_sql


def test_parse_valid_sql():
    parsed = parse_sql("SELECT * FROM Sheet1")
    assert parsed["action"] == "SELECT"
    assert parsed["table"] == "Sheet1"
    assert parsed["columns"] == ["*"]


def test_parse_invalid_sql():
    with pytest.raises(ValueError):
        parse_sql("INVALID SQL")


def test_parse_sql_with_where():
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = 1")
    assert parsed["where"]["column"] == "id"
    assert parsed["where"]["operator"] == "="
    assert parsed["where"]["value"] == 1


def test_parse_insert_with_columns_and_values():
    parsed = parse_sql("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')")
    assert parsed["action"] == "INSERT"
    assert parsed["table"] == "Sheet1"
    assert parsed["columns"] == ["id", "name"]
    assert parsed["values"] == [1, "Alice"]


def test_parse_insert_with_params():
    parsed = parse_sql(
        "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
        (2, "Bob"),
    )
    assert parsed["values"] == [2, "Bob"]


def test_parse_create_and_drop():
    parsed = parse_sql("CREATE TABLE Foo (id INT, name TEXT)")
    assert parsed["action"] == "CREATE"
    assert parsed["columns"] == ["id", "name"]

    parsed = parse_sql("DROP TABLE Foo")
    assert parsed["action"] == "DROP"
    assert parsed["table"] == "Foo"
