from __future__ import annotations

import pytest
from excel_dbapi.exceptions import DatabaseError
from typing import Any

from excel_dbapi.parser import parse_sql


VALID_CASES: list[tuple[str, dict[str, Any]]] = [
    (
        "UPDATE Users SET name = 'Bob' WHERE id = 1",
        {
            "action": "UPDATE",
            "table": "Users",
            "set": [{"column": "name", "value": {"type": "literal", "value": "Bob"}}],
            "where": {
                "conditions": [{"column": "id", "operator": "=", "value": 1}],
                "conjunctions": [],
            },
            "params": None,
        },
    ),
    (
        "UPDATE Users SET score = ?, note = ? WHERE id IN (?, ?)",
        {
            "action": "UPDATE",
            "table": "Users",
            "set": [
                {"column": "score", "value": {"type": "literal", "value": 9}},
                {
                    "column": "note",
                    "value": {"type": "literal", "value": "done"},
                },
            ],
            "where": {
                "conditions": [{"column": "id", "operator": "IN", "value": (1, 2)}],
                "conjunctions": [],
            },
            "params": (9, "done", 1, 2),
        },
    ),
    (
        "UPDATE Users SET note = 'x,y'",
        {
            "action": "UPDATE",
            "table": "Users",
            "set": [{"column": "note", "value": {"type": "literal", "value": "x,y"}}],
            "where": None,
            "params": None,
        },
    ),
]


INVALID_CASES: list[tuple[str, type[Exception], str]] = [
    ("UPDATE SET a = 1", DatabaseError, "Invalid UPDATE format"),
    ("UPDATE Users a = 1", DatabaseError, "Invalid UPDATE format"),
    ("UPDATE Users SET a", DatabaseError, "Invalid UPDATE format"),
    (
        "UPDATE Users SET a = ? WHERE b = ?",
        DatabaseError,
        "Missing parameters for placeholders",
    ),
    (
        "UPDATE Users SET a = ? WHERE b = ?",
        DatabaseError,
        "Not enough parameters for placeholders",
    ),
]


@pytest.mark.parametrize(("sql", "expected"), VALID_CASES)
def test_valid_parse(sql: str, expected: dict[str, Any]) -> None:
    params = expected["params"]
    assert parse_sql(sql, params) == expected


@pytest.mark.parametrize(
    ("sql", "exc_class", "msg", "params"),
    [
        (INVALID_CASES[0][0], INVALID_CASES[0][1], INVALID_CASES[0][2], None),
        (INVALID_CASES[1][0], INVALID_CASES[1][1], INVALID_CASES[1][2], None),
        (INVALID_CASES[2][0], INVALID_CASES[2][1], INVALID_CASES[2][2], None),
        (INVALID_CASES[3][0], INVALID_CASES[3][1], INVALID_CASES[3][2], None),
        (INVALID_CASES[4][0], INVALID_CASES[4][1], INVALID_CASES[4][2], (1,)),
    ],
)
def test_invalid_parse(
    sql: str,
    exc_class: type[Exception],
    msg: str,
    params: tuple[int, ...] | None,
) -> None:
    with pytest.raises(exc_class, match=msg):
        parse_sql(sql, params)
