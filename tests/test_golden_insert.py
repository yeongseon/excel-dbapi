from __future__ import annotations

import pytest
from typing import Any

from excel_dbapi.parser import parse_sql


VALID_CASES: list[tuple[str, dict[str, Any]]] = [
    (
        "INSERT INTO Users VALUES (1, 'Alice')",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": None,
            "values": [1, "Alice"],
            "params": None,
        },
    ),
    (
        "INSERT INTO Users (id, name) VALUES (2, 'Bob')",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": ["id", "name"],
            "values": [2, "Bob"],
            "params": None,
        },
    ),
    (
        "INSERT INTO Users (id, note) VALUES (?, ?)",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": ["id", "note"],
            "values": [3, "quoted, value"],
            "params": (3, "quoted, value"),
        },
    ),
    (
        'INSERT INTO Logs VALUES (4, "say ""hello""", NULL, 1.25)',
        {
            "action": "INSERT",
            "table": "Logs",
            "columns": None,
            "values": [4, 'say "hello"', None, 1.25],
            "params": None,
        },
    ),
]


INVALID_CASES: list[tuple[str, type[Exception], str]] = [
    ("INSERT Users VALUES (1)", ValueError, "Invalid INSERT format"),
    ("INSERT INTO Users (id)", ValueError, "Invalid INSERT format"),
    (
        "INSERT INTO Users (, ) VALUES (1)",
        ValueError,
        "Invalid column list",
    ),
    (
        "INSERT INTO Users (id, name) VALUES (?, ?)",
        ValueError,
        "Missing parameters for placeholders",
    ),
    (
        "INSERT INTO Users (id) VALUES (?)",
        ValueError,
        "Not enough parameters for placeholders",
    ),
    (
        "INSERT INTO Users (id) VALUES (1)",
        ValueError,
        "Too many parameters for placeholders",
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
        (INVALID_CASES[4][0], INVALID_CASES[4][1], INVALID_CASES[4][2], ()),
        (INVALID_CASES[5][0], INVALID_CASES[5][1], INVALID_CASES[5][2], (1,)),
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
