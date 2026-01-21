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
    assert parsed["where"]["conditions"][0]["column"] == "id"
    assert parsed["where"]["conditions"][0]["operator"] == "="
    assert parsed["where"]["conditions"][0]["value"] == 1


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


def test_parse_update_and_delete():
    parsed = parse_sql("UPDATE Sheet1 SET name = 'Bob' WHERE id = 1")
    assert parsed["action"] == "UPDATE"
    assert parsed["set"][0]["column"] == "name"
    assert parsed["set"][0]["value"] == "Bob"
    assert parsed["where"]["conditions"][0]["value"] == 1

    parsed = parse_sql("DELETE FROM Sheet1 WHERE id = 2")
    assert parsed["action"] == "DELETE"
    assert parsed["where"]["conditions"][0]["value"] == 2


def test_parse_select_with_order_limit_and_conditions():
    parsed = parse_sql("SELECT id, name FROM Sheet1 WHERE id >= 1 AND name = 'Alice' ORDER BY id DESC LIMIT 1")
    assert parsed["order_by"]["column"] == "id"
    assert parsed["order_by"]["direction"] == "DESC"
    assert parsed["limit"] == 1
    assert parsed["where"]["conditions"][0]["operator"] == ">="
