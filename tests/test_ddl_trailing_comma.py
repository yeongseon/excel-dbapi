"""Tests for DDL CREATE TABLE trailing comma rejection (issue #95)."""

from __future__ import annotations

import pytest

from excel_dbapi.parser.ddl import _parse_create


def test_trailing_comma_rejected() -> None:
    """CREATE TABLE t (a,) should raise ValueError."""
    with pytest.raises(ValueError, match="empty column definition"):
        _parse_create("CREATE TABLE t (a,)")


def test_trailing_comma_with_type_rejected() -> None:
    """CREATE TABLE t (a TEXT,) should raise ValueError."""
    with pytest.raises(ValueError, match="empty column definition"):
        _parse_create("CREATE TABLE t (a TEXT,)")


def test_multiple_trailing_commas_rejected() -> None:
    with pytest.raises(ValueError, match="empty column definition"):
        _parse_create("CREATE TABLE t (a,,)")


def test_leading_comma_rejected() -> None:
    with pytest.raises(ValueError, match="empty column definition"):
        _parse_create("CREATE TABLE t (,a)")


def test_valid_create_still_works() -> None:
    result = _parse_create("CREATE TABLE t (a, b TEXT)")
    assert result["columns"] == ["a", "b"]
    assert result["column_definitions"][0]["type_name"] == "TEXT"
    assert result["column_definitions"][1]["type_name"] == "TEXT"
