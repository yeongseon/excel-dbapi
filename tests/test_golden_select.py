from __future__ import annotations

import pytest
from typing import Any

from excel_dbapi.parser import parse_sql


VALID_CASES: list[tuple[str, dict[str, Any]]] = [
    (
        "SELECT * FROM Users",
        {
            "action": "SELECT",
            "columns": ["*"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": None,
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": None,
            "offset": None,
            "distinct": False,
            "params": None,
        },
    ),
    (
        "SELECT id, name FROM Users WHERE id IN (1, 2, 3)",
        {
            "action": "SELECT",
            "columns": ["id", "name"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": {
                "conditions": [{"column": "id", "operator": "IN", "value": (1, 2, 3)}],
                "conjunctions": [],
            },
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": None,
            "offset": None,
            "distinct": False,
            "params": None,
        },
    ),
    (
        "SELECT id FROM Users WHERE name = 'it''s' ORDER BY id DESC LIMIT 2",
        {
            "action": "SELECT",
            "columns": ["id"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": {
                "conditions": [{"column": "name", "operator": "=", "value": "it's"}],
                "conjunctions": [],
            },
            "group_by": None,
            "having": None,
            "order_by": {"column": "id", "direction": "DESC"},
            "limit": 2,
            "offset": None,
            "distinct": False,
            "params": None,
        },
    ),
    (
        'SELECT id FROM Users WHERE note = "say ""hello""" LIMIT ?',
        {
            "action": "SELECT",
            "columns": ["id"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": {
                "conditions": [
                    {"column": "note", "operator": "=", "value": 'say "hello"'}
                ],
                "conjunctions": [],
            },
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": 1,
            "offset": None,
            "distinct": False,
            "params": (1,),
        },
    ),
    (
        "SELECT id FROM Users WHERE score BETWEEN ? AND ? OR deleted IS NULL LIMIT ?",
        {
            "action": "SELECT",
            "columns": ["id"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": {
                "conditions": [
                    {"column": "score", "operator": "BETWEEN", "value": (1, 9)},
                    {"column": "deleted", "operator": "IS", "value": None},
                ],
                "conjunctions": ["OR"],
            },
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": 5,
            "offset": None,
            "distinct": False,
            "params": (1, 9, 5),
        },
    ),
    (
        "SELECT DISTINCT id, name FROM Users",
        {
            "action": "SELECT",
            "columns": ["id", "name"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": None,
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": None,
            "offset": None,
            "distinct": True,
            "params": None,
        },
    ),
    (
        "SELECT id FROM Users LIMIT 10 OFFSET 5",
        {
            "action": "SELECT",
            "columns": ["id"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": None,
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": 10,
            "offset": 5,
            "distinct": False,
            "params": None,
        },
    ),
    (
        "SELECT id FROM Users WHERE score >= ? LIMIT ? OFFSET ?",
        {
            "action": "SELECT",
            "columns": ["id"],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": {
                "conditions": [{"column": "score", "operator": ">=", "value": 50}],
                "conjunctions": [],
            },
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": 3,
            "offset": 1,
            "distinct": False,
            "params": (50, 3, 1),
        },
    ),
    (
        "SELECT COUNT(*) FROM Users",
        {
            "action": "SELECT",
            "columns": [{"type": "aggregate", "func": "COUNT", "arg": "*"}],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": None,
            "group_by": None,
            "having": None,
            "order_by": None,
            "limit": None,
            "offset": None,
            "distinct": False,
            "params": None,
        },
    ),
    (
        "SELECT name, COUNT(*) FROM Users GROUP BY name",
        {
            "action": "SELECT",
            "columns": ["name", {"type": "aggregate", "func": "COUNT", "arg": "*"}],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": None,
            "group_by": ["name"],
            "having": None,
            "order_by": None,
            "limit": None,
            "offset": None,
            "distinct": False,
            "params": None,
        },
    ),
    (
        "SELECT name, SUM(score) FROM Users GROUP BY name HAVING SUM(score) > 100",
        {
            "action": "SELECT",
            "columns": [
                "name",
                {"type": "aggregate", "func": "SUM", "arg": "score"},
            ],
            "table": "Users",
            "from": {"table": "Users", "alias": None, "ref": "Users"},
            "joins": None,
            "where": None,
            "group_by": ["name"],
            "having": {
                "conditions": [
                    {"column": "SUM(score)", "operator": ">", "value": 100}
                ],
                "conjunctions": [],
            },
            "order_by": None,
            "limit": None,
            "offset": None,
            "distinct": False,
            "params": None,
        },
    ),
]


INVALID_CASES: list[tuple[str, type[Exception], str]] = [
    ("SELECT , FROM Users", ValueError, "Invalid column list"),
    ("SELECT id FROM", ValueError, "Invalid SQL query format"),
    (
        "SELECT * FROM Users LIMIT WHERE id = 1",
        ValueError,
        "LIMIT cannot appear before WHERE",
    ),
    ("SELECT * FROM Users LIMIT ?", ValueError, "Missing parameters for placeholders"),
    (
        "SELECT * FROM Users ORDER BY  LIMIT 1",
        ValueError,
        "Invalid ORDER BY clause format",
    ),
    (
        "SELECT * FROM Users WHERE a = 1 XOR b = 2",
        ValueError,
        "Invalid WHERE clause format",
    ),
    (
        "SELECT * FROM Users WHERE a BETWEEN 1 2",
        ValueError,
        "Invalid WHERE clause format",
    ),
    (
        "SELECT * FROM Users WHERE a IN (",
        ValueError,
        "expected '\\)' in IN clause",
    ),
    (
        "SELECT * FROM Users WHERE a IN 1,2)",
        ValueError,
        "malformed IN clause",
    ),
    (
        "SELECT * FROM Users WHERE a IS MAYBE",
        ValueError,
        "expected NULL or NOT after IS",
    ),
    (
        "SELECT * FROM Users WHERE a IS NOT MAYBE",
        ValueError,
        "expected NULL after IS NOT",
    ),
    (
        "SELECT * FROM Users HAVING COUNT(*) > 1",
        ValueError,
        "HAVING requires GROUP BY",
    ),
    (
        "SELECT * FROM Users GROUP BY name WHERE id = 1",
        ValueError,
        "GROUP BY cannot appear before WHERE",
    ),
]


@pytest.mark.parametrize(("sql", "expected"), VALID_CASES)
def test_valid_parse(sql: str, expected: dict[str, Any]) -> None:
    params = expected["params"]
    assert parse_sql(sql, params) == expected


@pytest.mark.parametrize(("sql", "exc_class", "msg"), INVALID_CASES)
def test_invalid_parse(sql: str, exc_class: type[Exception], msg: str) -> None:
    with pytest.raises(exc_class, match=msg):
        parse_sql(sql)
