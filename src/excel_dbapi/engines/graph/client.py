"""Synchronous HTTP client for Microsoft Graph API with retry logic."""

from __future__ import annotations

import time
from typing import Any

import httpx

from ...exceptions import InterfaceError, OperationalError
from .auth import TokenProvider

_BASE_URL = "https://graph.microsoft.com/v1.0"
_RETRYABLE = frozenset({429, 503, 504})
_DEFAULT_MAX_RETRIES = 3
_DEFAULT_BACKOFF_FACTOR = 0.5  # seconds
_MAX_RETRY_AFTER = 60.0  # cap Retry-After to prevent excessive waits
_SAFE_METHODS = frozenset({"GET", "HEAD", "OPTIONS"})


def _parse_retry_after(value: str | None) -> float | None:
    """Parse a Retry-After header value (seconds or HTTP-date)."""
    if value is None:
        return None
    try:
        seconds = float(value)
        return min(seconds, _MAX_RETRY_AFTER)
    except ValueError:
        pass
    # HTTP-date format — fall back to default backoff
    return None


class GraphClient:
    """Thin synchronous wrapper around ``httpx.Client`` for Graph API calls.

    Features:
    - Bearer token injection via ``TokenProvider``
    - Workbook session header injection
    - Retry with exponential back-off on 429/503/504 (safe methods only)
    - Exception translation to DB-API ``OperationalError``
    """

    def __init__(
        self,
        token_provider: TokenProvider,
        *,
        transport: httpx.BaseTransport | None = None,
        timeout: float = 30.0,
        max_retries: int = _DEFAULT_MAX_RETRIES,
        backoff_factor: float = _DEFAULT_BACKOFF_FACTOR,
    ) -> None:
        kwargs: dict[str, Any] = {
            "base_url": _BASE_URL,
            "timeout": timeout,
        }
        if transport is not None:
            kwargs["transport"] = transport
        self._http = httpx.Client(**kwargs)
        self._token_provider = token_provider
        self._session_id: str | None = None
        self._max_retries = max(0, max_retries)
        self._backoff_factor = max(0.0, backoff_factor)

    # -- session management --------------------------------------------------

    @property
    def session_id(self) -> str | None:
        return self._session_id

    @session_id.setter
    def session_id(self, value: str | None) -> None:
        self._session_id = value

    # -- public request helpers ----------------------------------------------

    def get(self, path: str, **kwargs: Any) -> httpx.Response:
        return self._request("GET", path, **kwargs)

    def post(self, path: str, **kwargs: Any) -> httpx.Response:
        return self._request("POST", path, **kwargs)

    def patch(self, path: str, **kwargs: Any) -> httpx.Response:
        return self._request("PATCH", path, **kwargs)

    def delete(self, path: str, **kwargs: Any) -> httpx.Response:
        return self._request("DELETE", path, **kwargs)

    def close(self) -> None:
        self._http.close()

    # -- internals -----------------------------------------------------------

    def _build_headers(self) -> dict[str, str]:
        headers: dict[str, str] = {
            "Authorization": f"Bearer {self._token_provider.get_token()}",
            "Content-Type": "application/json",
        }
        if self._session_id is not None:
            headers["workbook-session-id"] = self._session_id
        return headers

    def _is_retryable(self, method: str) -> bool:
        """Only retry safe (idempotent read) methods automatically."""
        return method.upper() in _SAFE_METHODS

    @staticmethod
    def _format_error_message(status_code: int, message: str, body: str) -> str:
        if body:
            return f"Graph API error {status_code}: {message}. Response body: {body}"
        return f"Graph API error {status_code}: {message}"

    def _request(self, method: str, path: str, **kwargs: Any) -> httpx.Response:
        headers = {**self._build_headers(), **kwargs.pop("headers", {})}
        can_retry = self._is_retryable(method)
        last_exc: Exception | None = None
        max_attempts = (self._max_retries + 1) if can_retry else 1

        for attempt in range(max_attempts):
            try:
                resp = self._http.request(method, path, headers=headers, **kwargs)
            except httpx.TransportError as exc:
                last_exc = exc
                if can_retry and attempt < self._max_retries:
                    time.sleep(self._backoff_factor * (2**attempt))
                    continue
                raise OperationalError(f"Graph API request failed: {exc}") from exc

            if resp.status_code not in _RETRYABLE:
                try:
                    resp.raise_for_status()
                except httpx.HTTPStatusError as exc:
                    if resp.status_code == 401:
                        raise InterfaceError(
                            self._format_error_message(
                                401,
                                "Authentication expired or invalid. Re-authenticate and retry.",
                                resp.text,
                            )
                        ) from exc
                    if resp.status_code == 403:
                        raise InterfaceError(
                            self._format_error_message(
                                403,
                                "Insufficient permissions to access workbook. Check Graph API scopes: Files.ReadWrite.All",
                                resp.text,
                            )
                        ) from exc
                    if resp.status_code == 404:
                        raise OperationalError(
                            self._format_error_message(
                                404,
                                "Workbook not found. Check drive_id and item_id.",
                                resp.text,
                            )
                        ) from exc
                    raise OperationalError(
                        f"Graph API error {resp.status_code}: {resp.text}"
                    ) from exc
                return resp

            # Retryable status — only retry safe methods
            if not can_retry:
                raise OperationalError(
                    f"Graph API error {resp.status_code} on {method} (not retried): {resp.text}"
                )

            retry_after = _parse_retry_after(resp.headers.get("Retry-After"))
            wait = (
                retry_after
                if retry_after is not None
                else self._backoff_factor * (2**attempt)
            )
            if attempt < self._max_retries:
                time.sleep(wait)
            else:
                raise OperationalError(
                    f"Graph API error {resp.status_code} after {self._max_retries} retries"
                )

        # Should not reach here, but satisfy type checker
        raise last_exc or RuntimeError("Unexpected retry loop exit")  # pragma: no cover
