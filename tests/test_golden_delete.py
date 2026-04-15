from __future__ import annotations

import pytest
from excel_dbapi.exceptions import DatabaseError
from typing import Any

from excel_dbapi.parser import parse_sql


VALID_CASES: list[tuple[str, dict[str, Any]]] = [
    (
        "DELETE FROM Users",
        {
            "action": "DELETE",
            "table": "Users",
            "where": None,
            "params": None,
        },
    ),
    (
        "DELETE FROM Users WHERE id = 1",
        {
            "action": "DELETE",
            "table": "Users",
            "where": {
                "conditions": [{"column": "id", "operator": "=", "value": 1}],
                "conjunctions": [],
            },
            "params": None,
        },
    ),
    (
        "DELETE FROM Users WHERE id = ? OR deleted IS NOT NULL",
        {
            "action": "DELETE",
            "table": "Users",
            "where": {
                "conditions": [
                    {"column": "id", "operator": "=", "value": 9},
                    {"column": "deleted", "operator": "IS NOT", "value": None},
                ],
                "conjunctions": ["OR"],
            },
            "params": (9,),
        },
    ),
]


INVALID_CASES: list[tuple[str, type[Exception], str]] = [
    ("DELETE Users", DatabaseError, "Invalid DELETE format"),
    ("DELETE FROM Users ORDER BY id", DatabaseError, "Invalid DELETE format"),
    (
        "DELETE FROM Users WHERE id = ?",
        DatabaseError,
        "Missing parameters for placeholders",
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
