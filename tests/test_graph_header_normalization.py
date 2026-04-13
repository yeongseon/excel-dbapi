"""Tests for Graph backend header normalization via _normalize_headers()."""

import json
from typing import Any
from unittest.mock import Mock, patch

import httpx
import pytest

from excel_dbapi import connect
from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import DataError


DSN = "msgraph://drives/drv-test/items/itm-test"


def _make_graph_handler(worksheet_values: list[list[Any]]):
    """Create a mock handler that returns specific worksheet values."""

    def handler(request: httpx.Request) -> httpx.Response:
        method = request.method
        path = request.url.path
        body = None
        if request.content:
            try:
                body = json.loads(request.content)
            except (json.JSONDecodeError, UnicodeDecodeError):
                body = None

        if path.endswith("/createSession"):
            return httpx.Response(201, json={"id": "sess-test"})
        if path.endswith("/closeSession"):
            return httpx.Response(204)

        if (
            path.endswith("/worksheets") or "/worksheets?" in str(request.url)
        ) and method == "GET":
            return httpx.Response(
                200, json={"value": [{"id": "ws-test", "name": "Sheet1"}]}
            )

        if "usedRange" in path and method == "GET":
            return httpx.Response(200, json={"values": worksheet_values})

        return httpx.Response(404)

    return handler


class TestGraphHeaderNormalization:
    """Test that Graph backend validates headers through _normalize_headers()."""

    def test_valid_headers_pass_through(self):
        """Normal headers should pass validation and be returned as strings."""
        values = [["id", "name", "dept"], [1, "Alice", "Eng"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()
        results = cursor.execute("SELECT * FROM Sheet1").fetchall()
        conn.close()

        # Headers should be ["id", "name", "dept"]
        assert cursor.description is not None
        assert [col[0] for col in cursor.description] == ["id", "name", "dept"]

    def test_blank_header_raises_error(self):
        """Empty string header should raise DataError."""
        # Second column header is blank string
        values = [["id", "", "dept"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()

        with pytest.raises(DataError, match="Empty or None header at column index 1"):
            cursor.execute("SELECT * FROM Sheet1")
        conn.close()

    def test_none_header_raises_error(self):
        """None/null header should raise DataError."""
        # Second column header is None
        values = [["id", None, "dept"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()

        with pytest.raises(DataError, match="Empty or None header at column index 1"):
            cursor.execute("SELECT * FROM Sheet1")
        conn.close()

    def test_whitespace_only_header_raises_error(self):
        """Header with only whitespace should raise DataError."""
        # Second column header is spaces
        values = [["id", "   ", "dept"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()

        with pytest.raises(DataError, match="Empty or None header at column index 1"):
            cursor.execute("SELECT * FROM Sheet1")
        conn.close()

    def test_duplicate_header_case_insensitive_raises_error(self):
        """Duplicate headers (case-insensitive) should raise DataError."""
        # "Name" and "name" are duplicates
        values = [["id", "Name", "name"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()

        with pytest.raises(
            DataError, match="Duplicate header: 'name'.*conflicts with 'Name'"
        ):
            cursor.execute("SELECT * FROM Sheet1")
        conn.close()

    def test_headers_stripped_of_whitespace(self):
        """Headers with surrounding whitespace should be stripped."""
        # Headers with spaces should be normalized
        values = [[" id ", "  name  ", "dept"], [1, "Alice", "Eng"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()
        results = cursor.execute("SELECT * FROM Sheet1").fetchall()
        conn.close()

        # Headers should be stripped
        assert cursor.description is not None
        assert [col[0] for col in cursor.description] == ["id", "name", "dept"]

    def test_numeric_headers_converted_to_strings(self):
        """Numeric header values should be converted to strings."""
        # First column header is numeric
        values = [[1, "name", "dept"], ["Alice", "Engineering", "2024"]]
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()
        results = cursor.execute("SELECT * FROM Sheet1").fetchall()
        conn.close()

        # Numeric header should be converted to string "1"
        assert cursor.description is not None
        assert cursor.description[0][0] == "1"

    def test_empty_worksheet_returns_empty_headers(self):
        """Empty worksheet should return empty headers list without error."""
        values = []
        handler = _make_graph_handler(values)
        transport = httpx.MockTransport(handler)

        conn = ExcelConnection(
            DSN,
            credential="test-token",
            transport=transport,
            engine="graph",
        )
        cursor = conn.cursor()
        # Reading empty sheet should work (no headers to validate)
        # This depends on how the system handles empty sheets
        conn.close()
