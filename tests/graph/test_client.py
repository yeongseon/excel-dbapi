"""Tests for GraphClient retry, header injection, and exception translation."""

import json
from typing import Any

import httpx
import pytest

from excel_dbapi.engines.graph.auth import StaticTokenProvider
from excel_dbapi.engines.graph.client import GraphClient, _parse_retry_after
from excel_dbapi.exceptions import OperationalError


def _make_handler(responses: list[tuple[int, Any]]):
    """Return a handler that yields from a list of (status, body) pairs."""
    call_count = {"n": 0}

    def handler(request: httpx.Request) -> httpx.Response:
        idx = min(call_count["n"], len(responses) - 1)
        call_count["n"] += 1
        status, body = responses[idx]
        if body is None:
            return httpx.Response(status)
        return httpx.Response(status, json=body)

    handler.call_count = call_count  # type: ignore[attr-defined]
    return handler


class TestGraphClient:
    def _client(self, handler) -> GraphClient:
        transport = httpx.MockTransport(handler)
        return GraphClient(StaticTokenProvider("test-tok"), transport=transport)

    def test_get_includes_auth_header(self):
        captured: dict[str, Any] = {}

        def handler(request: httpx.Request) -> httpx.Response:
            captured["auth"] = request.headers.get("authorization")
            return httpx.Response(200, json={"ok": True})

        client = self._client(handler)
        client.get("/test")
        assert captured["auth"] == "Bearer test-tok"
        client.close()

    def test_session_id_header_injected(self):
        captured: dict[str, Any] = {}

        def handler(request: httpx.Request) -> httpx.Response:
            captured["session"] = request.headers.get("workbook-session-id")
            return httpx.Response(200, json={})

        client = self._client(handler)
        client.session_id = "sess-42"
        client.get("/test")
        assert captured["session"] == "sess-42"
        client.close()

    def test_no_session_header_when_none(self):
        captured: dict[str, Any] = {}

        def handler(request: httpx.Request) -> httpx.Response:
            captured["session"] = request.headers.get("workbook-session-id")
            return httpx.Response(200, json={})

        client = self._client(handler)
        client.get("/test")
        assert captured["session"] is None
        client.close()

    def test_retry_on_429_get(self):
        handler = _make_handler(
            [
                (429, None),
                (200, {"retried": True}),
            ]
        )
        client = self._client(handler)
        resp = client.get("/test")
        assert resp.json()["retried"] is True
        assert handler.call_count["n"] == 2
        client.close()

    def test_retry_on_503_get(self):
        handler = _make_handler(
            [
                (503, None),
                (503, None),
                (200, {"ok": True}),
            ]
        )
        client = self._client(handler)
        resp = client.get("/test")
        assert resp.json()["ok"] is True
        assert handler.call_count["n"] == 3
        client.close()

    def test_raises_operational_error_after_max_retries(self):
        handler = _make_handler(
            [
                (429, None),
                (429, None),
                (429, None),
                (429, None),  # 4th attempt = 3 retries exhausted
            ]
        )
        client = self._client(handler)
        with pytest.raises(OperationalError, match="after 3 retries"):
            client.get("/test")
        client.close()

    def test_post_sends_json(self):
        captured: dict[str, Any] = {}

        def handler(request: httpx.Request) -> httpx.Response:
            captured["body"] = json.loads(request.content)
            return httpx.Response(201, json={"id": "new"})

        client = self._client(handler)
        resp = client.post("/create", json={"key": "val"})
        assert resp.json()["id"] == "new"
        assert captured["body"] == {"key": "val"}
        client.close()

    def test_non_retryable_error_raises_operational_error(self):
        handler = _make_handler([(404, None)])
        client = self._client(handler)
        with pytest.raises(OperationalError, match="404"):
            client.get("/missing")
        assert handler.call_count["n"] == 1
        client.close()

    # -- Method-aware retry ---------------------------------------------------

    def test_post_not_retried_on_429(self):
        """POST is not a safe method — should NOT be retried."""
        handler = _make_handler(
            [
                (429, None),
                (200, {"ok": True}),  # should never reach here
            ]
        )
        client = self._client(handler)
        with pytest.raises(OperationalError, match="not retried"):
            client.post("/create", json={})
        assert handler.call_count["n"] == 1  # only 1 attempt
        client.close()

    def test_patch_not_retried_on_503(self):
        """PATCH is not a safe method — should NOT be retried."""
        handler = _make_handler(
            [
                (503, None),
                (200, {"ok": True}),
            ]
        )
        client = self._client(handler)
        with pytest.raises(OperationalError, match="not retried"):
            client.patch("/update", json={})
        assert handler.call_count["n"] == 1
        client.close()

    def test_delete_not_retried(self):
        """DELETE is not a safe method — should NOT be retried."""
        handler = _make_handler(
            [
                (503, None),
            ]
        )
        client = self._client(handler)
        with pytest.raises(OperationalError, match="not retried"):
            client.delete("/resource")
        assert handler.call_count["n"] == 1
        client.close()

    # -- Retry-After parsing ---------------------------------------------------

    def test_retry_after_header_numeric(self):
        """GET with Retry-After header (numeric) should be honoured."""
        responses = []

        def handler(request: httpx.Request) -> httpx.Response:
            responses.append(1)
            if len(responses) == 1:
                return httpx.Response(429, headers={"Retry-After": "0"})
            return httpx.Response(200, json={"ok": True})

        client = self._client(handler)
        resp = client.get("/test")
        assert resp.json()["ok"] is True
        client.close()


class TestParseRetryAfter:
    def test_none(self):
        assert _parse_retry_after(None) is None

    def test_numeric(self):
        assert _parse_retry_after("5") == 5.0

    def test_capped(self):
        assert _parse_retry_after("120") == 60.0  # capped at _MAX_RETRY_AFTER

    def test_http_date_returns_none(self):
        # HTTP-date format falls back to None (default backoff used)
        assert _parse_retry_after("Thu, 01 Jan 2099 00:00:00 GMT") is None
