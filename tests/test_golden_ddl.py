from __future__ import annotations

import pytest
from typing import Any

from excel_dbapi.parser import parse_sql
from excel_dbapi.exceptions import SqlParseError


VALID_CASES: list[tuple[str, dict[str, Any]]] = [
    (
        "CREATE TABLE Users (id INTEGER, name TEXT)",
        {
            "action": "CREATE",
            "table": "Users",
            "columns": ["id", "name"],
            "column_definitions": [
                {"name": "id", "type_name": "INTEGER"},
                {"name": "name", "type_name": "TEXT"},
            ],
            "params": None,
        },
    ),
    (
        "DROP TABLE Users",
        {
            "action": "DROP",
            "table": "Users",
            "params": None,
        },
    ),
]


INVALID_CASES: list[tuple[str, type[Exception], str]] = [
    ("CREATE TABLE", SqlParseError, "Invalid CREATE TABLE format"),
    ("CREATE TABLE Bad", SqlParseError, "Invalid CREATE TABLE format"),
    (
        "CREATE TABLE Bad (,)",
        SqlParseError,
        "Malformed column definitions: empty column definition found",
    ),
    (
        "CREATE TABLE Logs (message TEXT,)",
        SqlParseError,
        "Malformed column definitions: empty column definition found",
    ),
    ("DROP TABLE", SqlParseError, "Invalid DROP TABLE format"),
    ("DROP TABLE A B", SqlParseError, "Invalid DROP TABLE format"),
]


@pytest.mark.parametrize(("sql", "expected"), VALID_CASES)
def test_valid_parse(sql: str, expected: dict[str, Any]) -> None:
    assert parse_sql(sql) == expected


@pytest.mark.parametrize(("sql", "exc_class", "msg"), INVALID_CASES)
def test_invalid_parse(sql: str, exc_class: type[Exception], msg: str) -> None:
    with pytest.raises(exc_class, match=msg):
        parse_sql(sql)
