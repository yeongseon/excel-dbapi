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
            "values": [[1, "Alice"]],
            "params": None,
        },
    ),
    (
        "INSERT INTO Users (id, name) VALUES (2, 'Bob')",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": ["id", "name"],
            "values": [[2, "Bob"]],
            "params": None,
        },
    ),
    (
        "INSERT INTO Users (id, note) VALUES (?, ?)",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": ["id", "note"],
            "values": [[3, "quoted, value"]],
            "params": (3, "quoted, value"),
        },
    ),
    (
        'INSERT INTO Logs VALUES (4, "say ""hello""", NULL, 1.25)',
        {
            "action": "INSERT",
            "table": "Logs",
            "columns": None,
            "values": [[4, 'say "hello"', None, 1.25]],
            "params": None,
        },
    ),
    # Multi-row INSERT (2 rows)
    (
        "INSERT INTO Users VALUES (1, 'Alice'), (2, 'Bob')",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": None,
            "values": [[1, "Alice"], [2, "Bob"]],
            "params": None,
        },
    ),
    # Multi-row INSERT with columns
    (
        "INSERT INTO Users (id, name) VALUES (1, 'Alice'), (2, 'Bob')",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": ["id", "name"],
            "values": [[1, "Alice"], [2, "Bob"]],
            "params": None,
        },
    ),
    # Multi-row INSERT (3 rows)
    (
        "INSERT INTO Users VALUES (1, 'A'), (2, 'B'), (3, 'C')",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": None,
            "values": [[1, "A"], [2, "B"], [3, "C"]],
            "params": None,
        },
    ),
    # Multi-row INSERT with params
    (
        "INSERT INTO Users (id, name) VALUES (?, ?), (?, ?)",
        {
            "action": "INSERT",
            "table": "Users",
            "columns": ["id", "name"],
            "values": [[10, "X"], [20, "Y"]],
            "params": (10, "X", 20, "Y"),
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
    # INSERT without VALUES or SELECT
    ("INSERT INTO Users", ValueError, "Invalid INSERT format"),
    # VALUES with no tuples
    ("INSERT INTO Users VALUES", ValueError, "Invalid INSERT format"),
    # Unclosed tuple
    ("INSERT INTO Users VALUES (1, 'Alice'", ValueError, "Invalid INSERT format"),
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
        (INVALID_CASES[6][0], INVALID_CASES[6][1], INVALID_CASES[6][2], None),
        (INVALID_CASES[7][0], INVALID_CASES[7][1], INVALID_CASES[7][2], None),
        (INVALID_CASES[8][0], INVALID_CASES[8][1], INVALID_CASES[8][2], None),
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



def test_insert_select_golden() -> None:
    """INSERT...SELECT produces subquery values."""
    parsed = parse_sql("INSERT INTO Target SELECT id, name FROM Source")
    assert parsed["action"] == "INSERT"
    assert parsed["table"] == "Target"
    assert parsed["columns"] is None
    assert isinstance(parsed["values"], dict)
    assert parsed["values"]["type"] == "subquery"
    assert parsed["values"]["query"]["action"] == "SELECT"
    assert parsed["values"]["query"]["table"] == "Source"


def test_insert_select_with_columns_golden() -> None:
    """INSERT...SELECT with explicit column list."""
    parsed = parse_sql("INSERT INTO Target (id, name) SELECT id, name FROM Source")
    assert parsed["columns"] == ["id", "name"]
    assert parsed["values"]["type"] == "subquery"
    assert parsed["values"]["query"]["table"] == "Source"
