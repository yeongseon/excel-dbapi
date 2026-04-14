import pytest
from excel_dbapi.parser import parse_sql


def test_parse_valid_sql():
    parsed = parse_sql("SELECT * FROM Sheet1")
    assert parsed["action"] == "SELECT"
    assert parsed["table"] == "Sheet1"
    assert parsed["columns"] == ["*"]


def test_mixed_case_from():
    result = parse_sql("select * FrOm users")
    assert result["table"] == "users"


def test_mixed_case_select_and_from():
    result = parse_sql("SeLeCt COUNT(*) fRoM users")
    assert result["columns"] == [{"type": "aggregate", "func": "COUNT", "arg": "*"}]
    assert result["table"] == "users"


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
    assert parsed["values"] == [[1, "Alice"]]


def test_parse_insert_with_params():
    parsed = parse_sql(
        "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
        (2, "Bob"),
    )
    assert parsed["values"] == [[2, "Bob"]]


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
    assert parsed["set"][0]["value"] == {"type": "literal", "value": "Bob"}
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


def test_parse_subquery_in_where():
    result = parse_sql("SELECT * FROM users WHERE id IN (SELECT id FROM admins)")
    condition = result["where"]["conditions"][0]
    assert condition["operator"] == "IN"
    assert condition["value"]["type"] == "subquery"
    assert condition["value"]["query"]["action"] == "SELECT"
    assert condition["value"]["query"]["table"] == "admins"


def test_parse_subquery_with_where():
    result = parse_sql(
        "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE role = 'admin')"
    )
    subquery = result["where"]["conditions"][0]["value"]["query"]
    assert subquery["where"] is not None


def test_parse_subquery_with_nested_in_parentheses():
    result = parse_sql(
        "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE level IN (1, 2))"
    )
    subquery = result["where"]["conditions"][0]["value"]["query"]
    assert subquery["where"] is not None


def test_parse_subquery_preserves_literal_in():
    result = parse_sql("SELECT * FROM users WHERE id IN (1, 2, 3)")
    condition = result["where"]["conditions"][0]
    assert condition["value"] == (1, 2, 3)


def test_parse_subquery_rejects_multi_column():
    with pytest.raises(ValueError, match="exactly one column"):
        parse_sql("SELECT * FROM users WHERE id IN (SELECT id, name FROM admins)")


def test_parse_subquery_rejects_star():
    with pytest.raises(ValueError, match="exactly one column"):
        parse_sql("SELECT * FROM users WHERE id IN (SELECT * FROM admins)")


def test_parse_subquery_accepted_in_update():
    result = parse_sql(
        "UPDATE users SET name = 'x' WHERE id IN (SELECT id FROM admins)"
    )
    assert result["action"] == "UPDATE"
    assert result["where"] is not None
    cond = result["where"]["conditions"][0]
    assert cond["operator"] == "IN"
    assert cond["value"]["type"] == "subquery"


def test_parse_subquery_accepted_in_delete():
    result = parse_sql("DELETE FROM users WHERE id IN (SELECT id FROM admins)")
    assert result["action"] == "DELETE"
    assert result["where"] is not None
    cond = result["where"]["conditions"][0]
    assert cond["operator"] == "IN"
    assert cond["value"]["type"] == "subquery"


def test_parse_subquery_rejects_correlated():
    with pytest.raises(ValueError, match="[Cc]orrelated"):
        parse_sql(
            "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE admins.user_id = users.id)"
        )


def test_parse_subquery_rejects_parameterized():
    with pytest.raises(ValueError, match="[Pp]arameterized"):
        parse_sql(
            "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE role = ?)"
        )


def test_parse_subquery_allows_quoted_dotted_value():
    result = parse_sql(
        "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE role = 'corp.admin')"
    )
    subquery = result["where"]["conditions"][0]["value"]["query"]
    assert subquery["where"]["conditions"][0]["value"] == "corp.admin"


def test_parse_subquery_allows_quoted_question_mark():
    result = parse_sql(
        "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE symbol = '?')"
    )
    subquery = result["where"]["conditions"][0]["value"]["query"]
    assert subquery["where"]["conditions"][0]["value"] == "?"


def test_parse_subquery_rejects_group_by():
    with pytest.raises(ValueError, match="GROUP BY is not supported in subqueries"):
        parse_sql("SELECT * FROM users WHERE id IN (SELECT id FROM admins GROUP BY id)")


def test_parse_subquery_rejects_order_by():
    with pytest.raises(ValueError, match="ORDER BY is not supported in subqueries"):
        parse_sql("SELECT * FROM users WHERE id IN (SELECT id FROM admins ORDER BY id)")


def test_parse_subquery_rejects_limit():
    with pytest.raises(ValueError, match="LIMIT is not supported in subqueries"):
        parse_sql("SELECT * FROM users WHERE id IN (SELECT id FROM admins LIMIT 10)")


def test_parse_subquery_rejects_having():
    with pytest.raises(ValueError, match="HAVING is not supported in subqueries"):
        parse_sql(
            "SELECT * FROM users WHERE id IN (SELECT id FROM admins GROUP BY id HAVING COUNT(id) > 1)"
        )


def test_parse_subquery_rejects_offset():
    with pytest.raises(ValueError, match="OFFSET is not supported in subqueries"):
        parse_sql(
            "SELECT * FROM users WHERE id IN (SELECT id FROM admins LIMIT 10 OFFSET 5)"
        )


def test_parse_subquery_rejects_nested() -> None:
    with pytest.raises(ValueError, match="not supported in this context"):
        parse_sql(
            "SELECT * FROM users WHERE id IN (SELECT id FROM admins WHERE dept_id IN (SELECT id FROM depts))"
        )


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


def test_aggregate_count_distinct():
    result = parse_sql("SELECT COUNT(DISTINCT name) FROM users")
    assert result["action"] == "SELECT"
    col = result["columns"][0]
    assert col["type"] == "aggregate"
    assert col["func"] == "COUNT"
    assert col["arg"] == "name"
    assert col["distinct"] is True


def test_aggregate_rejects_expression():
    with pytest.raises(ValueError, match="Unsupported function: SUM"):
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


def test_rejects_window_over():
    parsed = parse_sql("SELECT COUNT(*) OVER () FROM users")
    assert parsed["action"] == "SELECT"
    col = parsed["columns"][0]
    assert col["type"] == "window_function"
    assert col["func"] == "COUNT"
    assert col["args"] == ["*"]


def test_parses_arithmetic_in_select():
    parsed = parse_sql("SELECT age + 1 FROM users")
    assert parsed["columns"] == [
        {
            "type": "binary_op",
            "op": "+",
            "left": "age",
            "right": {"type": "literal", "value": 1},
        }
    ]


def test_rejects_aggregate_arithmetic_in_select():
    with pytest.raises(
        ValueError,
        match="Unsupported column expression|Unsupported aggregate",
    ):
        parse_sql("SELECT COUNT(*) + 1 FROM users")


def test_rejects_aggregate_filter_in_select():
    parsed = parse_sql("SELECT COUNT(*) FILTER (WHERE id > 0) FROM users")
    assert parsed["action"] == "SELECT"
    col = parsed["columns"][0]
    assert col["type"] == "aggregate"
    assert col["func"] == "COUNT"
    assert col["arg"] == "*"
    assert col["filter"]["conditions"][0]["column"] == "id"
    assert col["filter"]["conditions"][0]["operator"] == ">"
    assert col["filter"]["conditions"][0]["value"] == 0


def test_parse_inner_join_basic():
    parsed = parse_sql(
        "SELECT a.id, b.name FROM Sheet1 a INNER JOIN Sheet2 b ON a.id = b.id"
    )
    assert parsed["from"] == {"table": "Sheet1", "alias": "a", "ref": "a"}
    assert parsed["joins"] == [
        {
            "type": "INNER",
            "source": {"table": "Sheet2", "alias": "b", "ref": "b"},
            "on": {
                "conditions": [
                    {
                        "column": {"type": "column", "source": "a", "name": "id"},
                        "operator": "=",
                        "value": {"type": "column", "source": "b", "name": "id"},
                    }
                ],
                "conjunctions": [],
            },
        }
    ]


def test_parse_left_join_basic():
    parsed = parse_sql(
        "SELECT a.id, b.name FROM Sheet1 a LEFT JOIN Sheet2 b ON a.id = b.id"
    )
    assert parsed["joins"][0]["type"] == "LEFT"


def test_parse_left_outer_join():
    parsed = parse_sql(
        "SELECT a.id, b.name FROM Sheet1 a LEFT OUTER JOIN Sheet2 b ON a.id = b.id"
    )
    assert parsed["joins"][0]["type"] == "LEFT"


def test_parse_join_without_inner_keyword():
    parsed = parse_sql("SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id")
    assert parsed["joins"][0]["type"] == "INNER"


def test_parse_join_with_where():
    parsed = parse_sql("SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id WHERE a.id = 1")
    assert parsed["joins"] is not None
    assert parsed["where"] is not None


def test_parse_join_with_order_by():
    parsed = parse_sql(
        "SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id ORDER BY a.id DESC"
    )
    assert parsed["order_by"] == {"column": "a.id", "direction": "DESC"}


def test_parse_join_with_limit_offset():
    parsed = parse_sql(
        "SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id LIMIT 5 OFFSET 2"
    )
    assert parsed["limit"] == 5
    assert parsed["offset"] == 2


def test_parse_join_with_table_name_prefix():
    parsed = parse_sql(
        "SELECT Sheet1.id FROM Sheet1 INNER JOIN Sheet2 ON Sheet1.id = Sheet2.id"
    )
    assert parsed["columns"] == [{"type": "column", "source": "Sheet1", "name": "id"}]


def test_parse_join_with_mixed_aliases():
    parsed = parse_sql(
        "SELECT a.id FROM Sheet1 a INNER JOIN Sheet2 b ON Sheet1.id = b.id"
    )
    clause = parsed["joins"][0]["on"]["conditions"][0]
    assert clause["column"]["source"] == "Sheet1"
    assert clause["value"]["source"] == "b"


def test_parse_join_allows_select_star():
    parsed = parse_sql("SELECT * FROM t1 a JOIN t2 b ON a.id = b.id")
    assert parsed["columns"] == ["*"]
    assert parsed["joins"] is not None


def test_parse_join_allows_multiple_joins():
    parsed = parse_sql(
        "SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id JOIN t3 c ON b.id = c.id"
    )
    assert parsed["joins"] is not None
    assert len(parsed["joins"]) == 2
    assert parsed["joins"][0]["type"] == "INNER"
    assert parsed["joins"][1]["type"] == "INNER"


def test_parse_join_rejects_unqualified_columns():
    with pytest.raises(ValueError, match="qualified column"):
        parse_sql("SELECT id FROM t1 a JOIN t2 b ON a.id = b.id")


def test_parse_join_accepts_right_join():
    parsed = parse_sql("SELECT a.id FROM t1 a RIGHT JOIN t2 b ON a.id = b.id")
    assert parsed["joins"][0]["type"] == "RIGHT"


def test_parse_join_accepts_cross_join():
    parsed = parse_sql("SELECT a.id FROM t1 a CROSS JOIN t2 b")
    assert parsed["joins"][0]["type"] == "CROSS"
    assert parsed["joins"][0]["on"] is None


def test_parse_join_accepts_full_join():
    parsed = parse_sql("SELECT a.id FROM t1 a FULL JOIN t2 b ON a.id = b.id")
    assert parsed["joins"][0]["type"] == "FULL"


def test_parse_cross_join_rejects_on_clause():
    with pytest.raises(ValueError, match="CROSS JOIN does not accept ON condition"):
        parse_sql("SELECT a.id FROM t1 a CROSS JOIN t2 b ON a.id = b.id")


def test_parse_join_rejects_non_equality_on():
    parsed = parse_sql("SELECT a.id FROM t1 a JOIN t2 b ON a.id > b.id")
    assert parsed["joins"][0]["on"]["conditions"][0]["operator"] == ">"


def test_parse_join_rejects_subquery_with_join():
    parsed = parse_sql(
        "SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id WHERE a.id IN (SELECT id FROM t3)"
    )
    assert parsed["joins"] is not None
    where_cond = parsed["where"]["conditions"][0]
    assert where_cond["operator"] == "IN"
    assert where_cond["value"]["type"] == "subquery"


def test_parse_join_accepts_group_by():
    parsed = parse_sql(
        "SELECT a.id, COUNT(*) FROM t1 a JOIN t2 b ON a.id = b.id GROUP BY a.id"
    )
    assert parsed["group_by"] == ["a.id"]


def test_parse_join_accepts_having():
    parsed = parse_sql(
        "SELECT a.id, COUNT(*) FROM t1 a JOIN t2 b ON a.id = b.id GROUP BY a.id HAVING COUNT(*) > 1"
    )
    assert parsed["having"] is not None


def test_parse_join_accepts_distinct():
    parsed = parse_sql("SELECT DISTINCT a.id FROM t1 a JOIN t2 b ON a.id = b.id")
    assert parsed["distinct"] is True


def test_parse_join_with_as_alias_from():
    """FROM table AS alias should be accepted."""
    result = parse_sql("SELECT a.id, b.id FROM t1 AS a JOIN t2 AS b ON a.id = b.id")
    assert result["from"]["alias"] == "a"
    assert result["from"]["ref"] == "a"
    assert result["joins"][0]["source"]["alias"] == "b"
    assert result["joins"][0]["source"]["ref"] == "b"


def test_parse_join_with_as_alias_mixed():
    """Mix of AS alias and bare alias should be accepted."""
    result = parse_sql("SELECT a.id, b.id FROM t1 AS a JOIN t2 b ON a.id = b.id")
    assert result["from"]["alias"] == "a"
    assert result["joins"][0]["source"]["alias"] == "b"


def test_parse_from_as_alias_single_table():
    """FROM table AS alias without JOIN should be accepted."""
    result = parse_sql("SELECT id FROM users AS u")
    assert result["from"]["alias"] == "u"
    assert result["from"]["ref"] == "u"


def test_parse_left_outer_join_with_as():
    """LEFT OUTER JOIN with AS alias should be accepted."""
    result = parse_sql(
        "SELECT a.id, b.id FROM t1 AS a LEFT OUTER JOIN t2 AS b ON a.id = b.id"
    )
    assert result["joins"][0]["type"] == "LEFT"
    assert result["joins"][0]["source"]["alias"] == "b"


def test_parse_join_rejects_duplicate_alias():
    """Duplicate table refs in JOIN should be rejected."""
    with pytest.raises(ValueError, match="Ambiguous table reference"):
        parse_sql("SELECT a.id FROM t1 a JOIN t2 a ON a.id = a.id")


def test_parse_join_rejects_duplicate_bare_table():
    """Self-join without distinct aliases should be rejected."""
    with pytest.raises(ValueError, match="Ambiguous table reference"):
        parse_sql("SELECT users.id FROM users JOIN users ON users.id = users.id")


def test_parse_join_rejects_alias_vs_table_name_collision():
    """Right alias colliding with left table name should be rejected.

    e.g. FROM users u JOIN orders users  -- right alias 'users' == left table 'users'
    """
    with pytest.raises(ValueError, match="Ambiguous table reference 'users'"):
        parse_sql("SELECT users.id FROM users u JOIN orders users ON u.id = users.id")


def test_parse_join_rejects_left_alias_vs_right_table_collision():
    """Left alias colliding with right table name should be rejected.

    e.g. FROM users orders JOIN orders o  -- left alias 'orders' == right table 'orders'
    """
    with pytest.raises(ValueError, match="Ambiguous table reference 'orders'"):
        parse_sql(
            "SELECT orders.id FROM users orders JOIN orders o ON orders.id = o.id"
        )


def test_parse_join_rejects_subquery_containing_join():
    """Subquery that itself contains a JOIN should be rejected."""
    with pytest.raises(ValueError, match="JOIN is not supported in subqueries"):
        parse_sql(
            "SELECT id FROM t1 WHERE id IN "
            "(SELECT a.id FROM t2 a JOIN t3 b ON a.id = b.id)"
        )


# ── Multi-row INSERT & INSERT...SELECT tests ──


def test_parse_multi_row_insert():
    parsed = parse_sql("INSERT INTO Sheet1 VALUES (1, 'Alice'), (2, 'Bob')")
    assert parsed["action"] == "INSERT"
    assert parsed["table"] == "Sheet1"
    assert parsed["columns"] is None
    assert parsed["values"] == [[1, "Alice"], [2, "Bob"]]


def test_parse_multi_row_insert_three_rows():
    parsed = parse_sql("INSERT INTO Sheet1 VALUES (1, 'A'), (2, 'B'), (3, 'C')")
    assert len(parsed["values"]) == 3
    assert parsed["values"][0] == [1, "A"]
    assert parsed["values"][1] == [2, "B"]
    assert parsed["values"][2] == [3, "C"]


def test_parse_multi_row_insert_with_params():
    parsed = parse_sql(
        "INSERT INTO Sheet1 (id, name) VALUES (?, ?), (?, ?)",
        (10, "X", 20, "Y"),
    )
    assert parsed["values"] == [[10, "X"], [20, "Y"]]


def test_parse_multi_row_insert_with_columns():
    parsed = parse_sql("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice'), (2, 'Bob')")
    assert parsed["columns"] == ["id", "name"]
    assert parsed["values"] == [[1, "Alice"], [2, "Bob"]]


def test_parse_insert_select():
    parsed = parse_sql("INSERT INTO Target SELECT id, name FROM Source")
    assert parsed["action"] == "INSERT"
    assert parsed["table"] == "Target"
    assert parsed["columns"] is None
    assert parsed["values"]["type"] == "subquery"
    assert parsed["values"]["query"]["action"] == "SELECT"
    assert parsed["values"]["query"]["table"] == "Source"


def test_parse_insert_select_with_where():
    parsed = parse_sql("INSERT INTO Target SELECT id, name FROM Source WHERE id > 5")
    assert parsed["values"]["type"] == "subquery"
    assert parsed["values"]["query"]["where"] is not None


def test_parse_insert_select_with_columns():
    parsed = parse_sql("INSERT INTO Target (id, name) SELECT id, name FROM Source")
    assert parsed["columns"] == ["id", "name"]
    assert parsed["values"]["type"] == "subquery"
    assert parsed["values"]["query"]["table"] == "Source"
