import pytest

from excel_dbapi.engine.parser import parse_sql


def test_parse_empty_query():
    with pytest.raises(ValueError):
        parse_sql("")


def test_parse_select_missing_from():
    with pytest.raises(ValueError):
        parse_sql("SELECT id Sheet1")


def test_parse_order_by_before_where():
    with pytest.raises(ValueError):
        parse_sql("SELECT * FROM Sheet1 ORDER BY id WHERE id = 1")


def test_parse_invalid_order_direction():
    with pytest.raises(ValueError):
        parse_sql("SELECT * FROM Sheet1 ORDER BY id DOWN")


def test_parse_invalid_limit():
    with pytest.raises(ValueError):
        parse_sql("SELECT * FROM Sheet1 LIMIT foo")


def test_parse_insert_missing_values():
    with pytest.raises(ValueError):
        parse_sql("INSERT INTO Sheet1 (id)")


def test_parse_insert_missing_params():
    with pytest.raises(ValueError):
        parse_sql("INSERT INTO Sheet1 (id) VALUES (?)")


def test_parse_update_missing_set():
    with pytest.raises(ValueError):
        parse_sql("UPDATE Sheet1 name = 'A'")


def test_parse_delete_missing_from():
    with pytest.raises(ValueError):
        parse_sql("DELETE Sheet1")


def test_parse_create_invalid_format():
    with pytest.raises(ValueError):
        parse_sql("CREATE TABLE Foo")


def test_parse_drop_invalid_format():
    with pytest.raises(ValueError):
        parse_sql("DROP Foo")
