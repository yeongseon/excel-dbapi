import pytest
from excel_dbapi.parser import parse_sql


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
    parsed = parse_sql(
        "SELECT id, name FROM Sheet1 WHERE id >= 1 AND name = 'Alice' ORDER BY id DESC LIMIT 1"
    )
    assert parsed["order_by"]["column"] == "id"
    assert parsed["order_by"]["direction"] == "DESC"
    assert parsed["limit"] == 1
    assert parsed["where"]["conditions"][0]["operator"] == ">="


def test_parse_update_with_or_where():
    parsed = parse_sql("UPDATE Sheet1 SET name = 'A' WHERE id = 1 OR id = 2")
    assert parsed["where"]["conjunctions"] == ["OR"]


def test_parse_select_distinct_variants():
    parsed = parse_sql("SELECT DISTINCT col1 FROM t")
    assert parsed["distinct"] is True
    assert parsed["columns"] == ["col1"]

    parsed = parse_sql("SELECT DISTINCT col1, col2 FROM t")
    assert parsed["distinct"] is True
    assert parsed["columns"] == ["col1", "col2"]

    parsed = parse_sql("SELECT DISTINCT * FROM t")
    assert parsed["distinct"] is True
    assert parsed["columns"] == ["*"]

    parsed = parse_sql("SELECT col1 FROM t")
    assert parsed["distinct"] is False


def test_parse_select_limit_offset_variants():
    parsed = parse_sql("SELECT * FROM t LIMIT 10 OFFSET 5")
    assert parsed["limit"] == 10
    assert parsed["offset"] == 5

    parsed = parse_sql("SELECT * FROM t OFFSET 5")
    assert parsed["limit"] is None
    assert parsed["offset"] == 5


def test_parse_select_limit_offset_param_binding_order():
    parsed = parse_sql("SELECT * FROM t LIMIT ? OFFSET ?", (10, 5))
    assert parsed["limit"] == 10
    assert parsed["offset"] == 5

    parsed = parse_sql("SELECT * FROM t WHERE x = ? LIMIT ? OFFSET ?", (1, 10, 5))
    assert parsed["where"]["conditions"][0]["value"] == 1
    assert parsed["limit"] == 10
    assert parsed["offset"] == 5


def test_parse_select_count_star_column():
    parsed = parse_sql("SELECT COUNT(*) FROM Sheet1")
    assert parsed["columns"] == [{"type": "aggregate", "func": "COUNT", "arg": "*"}]
    assert parsed["group_by"] is None
    assert parsed["having"] is None


def test_parse_select_with_group_by_and_count():
    parsed = parse_sql("SELECT name, COUNT(*) FROM Sheet1 GROUP BY name")
    assert parsed["columns"] == [
        "name",
        {"type": "aggregate", "func": "COUNT", "arg": "*"},
    ]
    assert parsed["group_by"] == ["name"]


def test_parse_select_with_having_aggregate_expression():
    parsed = parse_sql(
        "SELECT name, SUM(score) FROM Sheet1 GROUP BY name HAVING SUM(score) > 100"
    )
    assert parsed["group_by"] == ["name"]
    assert parsed["having"] == {
        "conditions": [{"column": "SUM(score)", "operator": ">", "value": 100}],
        "conjunctions": [],
    }


def test_parse_select_order_by_aggregate_expression():
    parsed = parse_sql(
        "SELECT name, COUNT(*) FROM Sheet1 GROUP BY name ORDER BY COUNT(*) DESC"
    )
    assert parsed["order_by"] == {"column": "COUNT(*)", "direction": "DESC"}


def test_parse_select_all_aggregate_functions():
    parsed = parse_sql(
        "SELECT COUNT(*), COUNT(score), SUM(score), AVG(score), MIN(score), MAX(score) FROM Sheet1"
    )
    assert parsed["columns"] == [
        {"type": "aggregate", "func": "COUNT", "arg": "*"},
        {"type": "aggregate", "func": "COUNT", "arg": "score"},
        {"type": "aggregate", "func": "SUM", "arg": "score"},
        {"type": "aggregate", "func": "AVG", "arg": "score"},
        {"type": "aggregate", "func": "MIN", "arg": "score"},
        {"type": "aggregate", "func": "MAX", "arg": "score"},
    ]


def test_parse_select_group_by_before_where_raises():
    with pytest.raises(ValueError):
        parse_sql("SELECT name, COUNT(*) FROM Sheet1 GROUP BY name WHERE name = 'A'")


def test_where_rejects_aggregate_count():
    with pytest.raises(
        ValueError,
        match="Aggregate functions are not allowed in WHERE",
    ):
        parse_sql("SELECT name FROM users WHERE COUNT(*) > 1")


def test_where_rejects_aggregate_sum():
    with pytest.raises(
        ValueError,
        match="Aggregate functions are not allowed in WHERE",
    ):
        parse_sql("SELECT name FROM users WHERE SUM(score) > 100")


def test_where_rejects_aggregate_avg():
    with pytest.raises(
        ValueError,
        match="Aggregate functions are not allowed in WHERE",
    ):
        parse_sql("SELECT name FROM users WHERE AVG(score) > 50")


def test_where_with_quoted_group_by_keyword():
    result = parse_sql("SELECT * FROM users WHERE note = 'x GROUP BY y'")
    assert result["action"] == "SELECT"
    assert result["where"] is not None
    assert result["group_by"] is None


def test_where_with_quoted_order_by_keyword():
    result = parse_sql("SELECT * FROM users WHERE note = 'x ORDER BY y'")
    assert result["action"] == "SELECT"
    assert result["where"] is not None
    assert result["order_by"] is None


def test_aggregate_rejects_distinct():
    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        parse_sql("SELECT COUNT(DISTINCT name) FROM users")


def test_aggregate_rejects_expression():
    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        parse_sql("SELECT SUM(age + 1) FROM users")


def test_aggregate_rejects_numeric_literal():
    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        parse_sql("SELECT COUNT(1) FROM users")


def test_aggregate_rejects_string_literal():
    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        parse_sql("SELECT COUNT('x') FROM users")


def test_aggregate_rejects_float_literal():
    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        parse_sql("SELECT SUM(3.14) FROM users")
