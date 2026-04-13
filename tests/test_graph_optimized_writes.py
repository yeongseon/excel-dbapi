from __future__ import annotations

import json
import re
from typing import Any

import httpx

from excel_dbapi.connection import ExcelConnection


DSN = "msgraph://drives/drv-opt/items/itm-opt"


def _col_index(letter: str) -> int:
    value = 0
    for char in letter:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def _parse_address(path: str) -> tuple[int, int, int, int]:
    match = re.search(r"range\(address='([A-Z]+)(\d+):([A-Z]+)(\d+)'\)", path)
    if match is None:
        raise ValueError(f"Invalid range address in path: {path}")
    start_col = _col_index(match.group(1))
    start_row = int(match.group(2)) - 1
    end_col = _col_index(match.group(3))
    end_row = int(match.group(4)) - 1
    return start_row, end_row, start_col, end_col


def _build_handler() -> tuple[httpx.MockTransport, dict[str, Any]]:
    state: dict[str, Any] = {
        "worksheets": {
            "ws-emp": {
                "name": "Employees",
                "values": [
                    ["id", "name", "dept"],
                    [1, "Alice", "Eng"],
                    [2, "Bob", "Sales"],
                    [3, "Carol", "Eng"],
                ],
            }
        },
        "requests": [],
    }

    def handler(request: httpx.Request) -> httpx.Response:
        method = request.method
        path = request.url.path
        body = None
        if request.content:
            body = json.loads(request.content)

        state["requests"].append((method, path, body))

        if path.endswith("/createSession"):
            return httpx.Response(201, json={"id": "sess-opt"})
        if path.endswith("/closeSession"):
            return httpx.Response(204)

        if (
            path.endswith("/worksheets") or "/worksheets?" in str(request.url)
        ) and method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {"id": ws_id, "name": ws_data["name"]}
                        for ws_id, ws_data in state["worksheets"].items()
                    ]
                },
            )

        if "usedRange" in path and method == "GET":
            for ws_id, ws_data in state["worksheets"].items():
                if ws_id in path:
                    return httpx.Response(200, json={"values": ws_data["values"]})
            return httpx.Response(200, json={"values": []})

        if "/range(" in path and method == "PATCH":
            if body is None:
                return httpx.Response(400, json={"error": "missing body"})
            start_row, end_row, start_col, end_col = _parse_address(path)
            values = body["values"]
            ws = state["worksheets"]["ws-emp"]
            sheet_values: list[list[Any]] = ws["values"]

            width = max((len(row) for row in sheet_values), default=0)
            width = max(width, end_col + 1)
            for row in sheet_values:
                if len(row) < width:
                    row.extend([None] * (width - len(row)))

            while len(sheet_values) <= end_row:
                sheet_values.append([None] * width)

            expected_rows = end_row - start_row + 1
            if len(values) != expected_rows:
                return httpx.Response(400, json={"error": "row mismatch"})

            for offset, patch_row in enumerate(values):
                row_index = start_row + offset
                current = sheet_values[row_index]
                update_row = list(patch_row)
                needed = end_col - start_col + 1
                if len(update_row) < needed:
                    update_row.extend([None] * (needed - len(update_row)))
                current[start_col : end_col + 1] = update_row[:needed]

            return httpx.Response(200, json={})

        if path.endswith("/delete") and method == "POST":
            start_row, end_row, _, _ = _parse_address(path)
            ws = state["worksheets"]["ws-emp"]
            sheet_values: list[list[Any]] = ws["values"]
            del sheet_values[start_row : end_row + 1]
            return httpx.Response(200, json={})

        if path.endswith("/clear") and method == "POST":
            return httpx.Response(200, json={})

        return httpx.Response(404)

    return httpx.MockTransport(handler), state


def _make_connection() -> tuple[ExcelConnection, dict[str, Any]]:
    transport, state = _build_handler()
    conn = ExcelConnection(DSN, credential="tok", transport=transport, readonly=False)
    return conn, state


def test_single_row_update_uses_targeted_patch() -> None:
    conn, state = _make_connection()
    cursor = conn.cursor()

    cursor.execute("UPDATE Employees SET dept = 'HR' WHERE id = 1")
    assert cursor.rowcount == 1

    patch_requests = [r for r in state["requests"] if r[0] == "PATCH"]
    assert len(patch_requests) == 1
    assert "A2:C2" in patch_requests[0][1]
    assert patch_requests[0][2]["values"] == [[1, "Alice", "HR"]]
    conn.close()


def test_multi_row_update_batches_contiguous_rows() -> None:
    conn, state = _make_connection()
    state["worksheets"]["ws-emp"]["values"].extend(
        [
            [4, "Dave", "Ops"],
            [5, "Eve", "Ops"],
        ]
    )
    cursor = conn.cursor()

    cursor.execute("UPDATE Employees SET dept = 'Ops' WHERE id <= 2")
    assert cursor.rowcount == 2

    patch_requests = [r for r in state["requests"] if r[0] == "PATCH"]
    assert len(patch_requests) == 1
    assert "A2:C3" in patch_requests[0][1]
    assert patch_requests[0][2]["values"] == [
        [1, "Alice", "Ops"],
        [2, "Bob", "Ops"],
    ]
    conn.close()


def test_delete_uses_row_delete_endpoint() -> None:
    conn, state = _make_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM Employees WHERE dept = 'Eng'")
    assert cursor.rowcount == 2

    delete_requests = [
        r for r in state["requests"] if r[0] == "POST" and r[1].endswith("/delete")
    ]
    assert len(delete_requests) == 2
    assert "A4:C4" in delete_requests[0][1]
    assert "A2:C2" in delete_requests[1][1]
    assert state["worksheets"]["ws-emp"]["values"] == [
        ["id", "name", "dept"],
        [2, "Bob", "Sales"],
    ]
    conn.close()


def test_large_update_falls_back_to_full_rewrite() -> None:
    conn, state = _make_connection()
    cursor = conn.cursor()

    cursor.execute("UPDATE Employees SET dept = 'All'")
    assert cursor.rowcount == 3

    patch_requests = [r for r in state["requests"] if r[0] == "PATCH"]
    assert len(patch_requests) == 1
    assert "A1:C4" in patch_requests[0][1]
    assert patch_requests[0][2]["values"] == [
        ["id", "name", "dept"],
        [1, "Alice", "All"],
        [2, "Bob", "All"],
        [3, "Carol", "All"],
    ]
    conn.close()


def test_data_integrity_after_optimized_update_and_delete() -> None:
    conn, state = _make_connection()
    cursor = conn.cursor()

    cursor.execute("UPDATE Employees SET dept = 'HR' WHERE id = 1")
    cursor.execute("DELETE FROM Employees WHERE id = 2")
    cursor.execute("SELECT id, name, dept FROM Employees ORDER BY id")

    rows = cursor.fetchall()
    assert rows == [(1, "Alice", "HR"), (3, "Carol", "Eng")]
    assert state["worksheets"]["ws-emp"]["values"] == [
        ["id", "name", "dept"],
        [1, "Alice", "HR"],
        [3, "Carol", "Eng"],
    ]
    conn.close()
