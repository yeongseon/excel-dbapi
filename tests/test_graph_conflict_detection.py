from typing import Any

import httpx
import pytest

from excel_dbapi.engines.graph.backend import GraphBackend
from excel_dbapi.exceptions import OperationalError

DSN = "msgraph://drives/drv-1/items/itm-1"


def _build_conflict_handler(
    *,
    fail_on_patch: bool = False,
    initial_etag: str = '"v1"',
    updated_etag: str = '"v2"',
) -> tuple[Any, dict[str, Any]]:
    state: dict[str, Any] = {
        "request_headers": [],
    }

    def handler(request: httpx.Request) -> httpx.Response:
        method = request.method
        path = request.url.path
        state["request_headers"].append((method, path, dict(request.headers)))

        if path.endswith("/createSession") and method == "POST":
            return httpx.Response(201, json={"id": "sess-conflict"})

        if path.endswith("/closeSession") and method == "POST":
            return httpx.Response(204)

        if path.endswith("/workbook") and method == "GET":
            return httpx.Response(200, json={"@odata.etag": initial_etag})

        if (
            path.endswith("/worksheets") or "/worksheets?" in str(request.url)
        ) and method == "GET":
            return httpx.Response(
                200, json={"value": [{"id": "ws-1", "name": "Users"}]}
            )

        if "usedRange" in path and method == "GET":
            return httpx.Response(200, json={"values": [["id"], [1], [2]]})

        if "/range(" in path and method == "PATCH":
            if fail_on_patch:
                return httpx.Response(
                    412, json={"error": {"code": "preconditionFailed"}}
                )
            return httpx.Response(200, json={}, headers={"ETag": updated_etag})

        return httpx.Response(404, json={"error": {"code": "notFound"}})

    return handler, state


def _make_backend(
    *, conflict_strategy: str = "fail", fail_on_patch: bool = False
) -> tuple[GraphBackend, dict[str, Any]]:
    handler, state = _build_conflict_handler(fail_on_patch=fail_on_patch)
    backend = GraphBackend(
        DSN,
        credential="test-token",
        transport=httpx.MockTransport(handler),
        readonly=False,
        conflict_strategy=conflict_strategy,
    )
    return backend, state


def _patch_headers(state: dict[str, Any]) -> list[dict[str, str]]:
    return [
        headers
        for method, path, headers in state["request_headers"]
        if method == "PATCH" and "/range(" in path
    ]


def test_etag_tracked_from_initial_open() -> None:
    backend, _ = _make_backend()
    backend.load()
    assert backend._etag == '"v1"'
    backend.close()


def test_if_match_header_sent_on_write() -> None:
    backend, state = _make_backend()
    backend.append_row("Users", [3])
    patch_headers = _patch_headers(state)
    assert len(patch_headers) == 1
    assert patch_headers[0].get("if-match") == '"v1"'
    backend.close()


def test_precondition_failed_raises_operational_error() -> None:
    backend, _ = _make_backend(fail_on_patch=True)
    with pytest.raises(
        OperationalError,
        match="Concurrent modification detected: workbook was modified by another session",
    ):
        backend.append_row("Users", [3])
    backend.close()


def test_etag_updated_after_successful_write() -> None:
    backend, _ = _make_backend()
    backend.append_row("Users", [3])
    assert backend._etag == '"v2"'
    backend.close()


def test_conflict_strategy_force_bypasses_if_match() -> None:
    backend, state = _make_backend(conflict_strategy="force")
    backend.append_row("Users", [3])
    patch_headers = _patch_headers(state)
    assert len(patch_headers) == 1
    assert patch_headers[0].get("if-match") is None
    backend.close()
