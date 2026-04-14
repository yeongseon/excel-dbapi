from __future__ import annotations

import json
import re
from typing import Any

import httpx
import pytest

from excel_dbapi.engines.base import TableData
from excel_dbapi.engines.graph.auth import StaticTokenProvider
from excel_dbapi.engines.graph.backend import GraphBackend
from excel_dbapi.engines.graph.client import GraphClient
from excel_dbapi.exceptions import OperationalError


DSN = "msgraph://drives/drv-stress/items/itm-stress"


def _column_index(label: str) -> int:
    value = 0
    for char in label:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def _parse_range_address(path: str) -> tuple[int, int, int, int]:
    match = re.search(r"range\(address='([A-Z]+)(\d+):([A-Z]+)(\d+)'\)", path)
    if match is None:
        raise ValueError(f"Invalid range address: {path}")
    start_col = _column_index(match.group(1))
    start_row = int(match.group(2)) - 1
    end_col = _column_index(match.group(3))
    end_row = int(match.group(4)) - 1
    return start_row, end_row, start_col, end_col


def _build_backend(
    *,
    conflict_strategy: str = "fail",
    conflict_on_if_match: bool = False,
) -> tuple[GraphBackend, dict[str, Any]]:
    state: dict[str, Any] = {
        "values": [["id", "name"], [1, "seed"]],
        "etag": '"v1"',
        "requests": [],
    }

    def handler(request: httpx.Request) -> httpx.Response:
        path = request.url.path
        method = request.method
        state["requests"].append((method, path, dict(request.headers), request.content))

        if path.endswith("/createSession") and method == "POST":
            return httpx.Response(201, json={"id": "sess-stress"})
        if path.endswith("/closeSession") and method == "POST":
            return httpx.Response(204)
        if path.endswith("/workbook") and method == "GET":
            return httpx.Response(200, json={"@odata.etag": state["etag"]})
        if (
            path.endswith("/worksheets") or "/worksheets?" in str(request.url)
        ) and method == "GET":
            return httpx.Response(
                200, json={"value": [{"id": "ws-1", "name": "Users"}]}
            )
        if "usedRange" in path and method == "GET":
            return httpx.Response(200, json={"values": state["values"]})

        if "/range(" in path and method == "PATCH":
            if conflict_on_if_match and request.headers.get("if-match"):
                return httpx.Response(
                    412, json={"error": {"code": "preconditionFailed"}}
                )

            start_row, end_row, start_col, end_col = _parse_range_address(path)
            payload = json.loads(request.content.decode("utf-8"))
            patch_values = payload["values"]

            width = max((len(row) for row in state["values"]), default=0)
            width = max(width, end_col + 1)
            for row in state["values"]:
                if len(row) < width:
                    row.extend([None] * (width - len(row)))

            while len(state["values"]) <= end_row:
                state["values"].append([None] * width)

            for offset, patch_row in enumerate(patch_values):
                row_index = start_row + offset
                normalized = list(patch_row)
                needed = end_col - start_col + 1
                if len(normalized) < needed:
                    normalized.extend([None] * (needed - len(normalized)))
                state["values"][row_index][start_col : end_col + 1] = normalized[
                    :needed
                ]

            current = int(state["etag"].strip('"v'))
            state["etag"] = f'"v{current + 1}"'
            return httpx.Response(200, json={}, headers={"ETag": state["etag"]})

        if path.endswith("/clear") and method == "POST":
            return httpx.Response(200, json={})
        if path.endswith("/delete") and method == "POST":
            return httpx.Response(200, json={})

        return httpx.Response(404, json={"error": {"code": "notFound"}})

    backend = GraphBackend(
        DSN,
        credential="token",
        transport=httpx.MockTransport(handler),
        readonly=False,
        conflict_strategy=conflict_strategy,
    )
    return backend, state


def test_graph_large_payload_write_stress() -> None:
    backend, state = _build_backend()
    rows = [[idx, f"user-{idx}"] for idx in range(1, 1201)]

    backend.write_sheet("Users", TableData(headers=["id", "name"], rows=rows))
    backend.close()

    patch_calls = [req for req in state["requests"] if req[0] == "PATCH"]
    assert patch_calls
    assert len(state["values"]) == 1201


def test_graph_multiple_sequential_writes_stress() -> None:
    backend, state = _build_backend()

    for idx in range(2, 102):
        backend.append_row("Users", [idx, f"u-{idx}"])
    backend.close()

    assert len(state["values"]) == 102
    patch_calls = [req for req in state["requests"] if req[0] == "PATCH"]
    assert len(patch_calls) == 100


def test_graph_retry_exhaustion_stress(monkeypatch: pytest.MonkeyPatch) -> None:
    sleeps: list[float] = []
    monkeypatch.setattr("excel_dbapi.engines.graph.client.time.sleep", sleeps.append)

    attempts = 0

    def handler(_: httpx.Request) -> httpx.Response:
        nonlocal attempts
        attempts += 1
        return httpx.Response(429, json={"error": "rate-limited"})

    client = GraphClient(
        StaticTokenProvider("token"),
        transport=httpx.MockTransport(handler),
        max_retries=2,
        backoff_factor=0.01,
    )

    with pytest.raises(OperationalError, match="after 2 retries"):
        client.get("/retry")
    client.close()

    assert attempts == 3
    assert sleeps == [0.01, 0.02]


def test_graph_etag_conflict_resolution_under_contention() -> None:
    backend, state = _build_backend(
        conflict_strategy="force",
        conflict_on_if_match=True,
    )

    for idx in range(2, 12):
        backend.append_row("Users", [idx, f"contended-{idx}"])
    backend.close()

    patch_calls = [req for req in state["requests"] if req[0] == "PATCH"]
    assert len(patch_calls) == 10
    assert all("if-match" not in headers for _, _, headers, _ in patch_calls)
    assert len(state["values"]) == 12
