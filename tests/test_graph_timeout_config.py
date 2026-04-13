from __future__ import annotations

import httpx
import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.engines.graph.auth import StaticTokenProvider
from excel_dbapi.engines.graph.client import GraphClient
from excel_dbapi.exceptions import InterfaceError, OperationalError


DSN = "msgraph://drives/drv-timeout/items/itm-timeout"


def test_custom_timeout_flows_from_connection_backend_options() -> None:
    captured_timeouts: list[float] = []

    def handler(request: httpx.Request) -> httpx.Response:
        timeout_ext = request.extensions.get("timeout")
        assert isinstance(timeout_ext, dict)
        captured_timeouts.append(float(timeout_ext["read"]))

        path = request.url.path
        if path.endswith("/createSession"):
            return httpx.Response(201, json={"id": "sess-timeout"})
        if path.endswith("/closeSession"):
            return httpx.Response(204)
        if path.endswith("/worksheets") or "/worksheets?" in str(request.url):
            return httpx.Response(200, json={"value": [{"id": "ws-1", "name": "Users"}]})
        if "usedRange" in path:
            return httpx.Response(200, json={"values": [["id", "name"], [1, "Ada"]]})
        return httpx.Response(404)

    transport = httpx.MockTransport(handler)
    conn = ExcelConnection(DSN, credential="token", transport=transport, timeout=7.5)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Users")
    assert cursor.fetchall() == [(1, "Ada")]
    conn.close()

    assert captured_timeouts
    assert all(value == 7.5 for value in captured_timeouts)


def test_custom_retry_count_limits_attempts() -> None:
    call_count = 0

    def handler(_: httpx.Request) -> httpx.Response:
        nonlocal call_count
        call_count += 1
        return httpx.Response(429, json={"error": "rate_limited"})

    client = GraphClient(
        StaticTokenProvider("test-token"),
        transport=httpx.MockTransport(handler),
        max_retries=1,
    )
    with pytest.raises(OperationalError, match="after 1 retries"):
        client.get("/retry")
    client.close()
    assert call_count == 2


def test_custom_backoff_factor_controls_sleep(monkeypatch: pytest.MonkeyPatch) -> None:
    sleeps: list[float] = []
    monkeypatch.setattr("excel_dbapi.engines.graph.client.time.sleep", sleeps.append)

    call_count = 0

    def handler(_: httpx.Request) -> httpx.Response:
        nonlocal call_count
        call_count += 1
        if call_count < 3:
            return httpx.Response(503, json={"error": "busy"})
        return httpx.Response(200, json={"ok": True})

    client = GraphClient(
        StaticTokenProvider("test-token"),
        transport=httpx.MockTransport(handler),
        max_retries=2,
        backoff_factor=0.25,
    )
    resp = client.get("/backoff")
    client.close()

    assert resp.status_code == 200
    assert sleeps == [0.25, 0.5]


@pytest.mark.parametrize(
    ("status_code", "expected_message"),
    [
        (
            401,
            "Authentication expired or invalid. Re-authenticate and retry.",
        ),
        (
            403,
            "Insufficient permissions to access workbook. Check Graph API scopes: Files.ReadWrite.All",
        ),
    ],
)
def test_auth_errors_raise_interface_error_with_diagnostics(
    status_code: int,
    expected_message: str,
) -> None:
    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(status_code, json={"error": {"message": "token problem"}})

    client = GraphClient(
        StaticTokenProvider("test-token"),
        transport=httpx.MockTransport(handler),
    )
    with pytest.raises(InterfaceError) as exc_info:
        client.get("/secure")
    client.close()

    err = str(exc_info.value)
    assert expected_message in err
    assert "token problem" in err


def test_default_values_are_backwards_compatible() -> None:
    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json={"ok": True})

    client = GraphClient(
        StaticTokenProvider("test-token"),
        transport=httpx.MockTransport(handler),
    )
    timeout = client._http.timeout
    client.close()

    assert timeout.connect == 30.0
    assert timeout.read == 30.0
    assert client._max_retries == 3
    assert client._backoff_factor == 0.5


def test_404_includes_workbook_guidance_and_body() -> None:
    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(404, json={"error": "missing workbook"})

    client = GraphClient(
        StaticTokenProvider("test-token"),
        transport=httpx.MockTransport(handler),
    )
    with pytest.raises(OperationalError) as exc_info:
        client.get("/missing")
    client.close()

    err = str(exc_info.value)
    assert "Workbook not found. Check drive_id and item_id." in err
    assert "missing workbook" in err
