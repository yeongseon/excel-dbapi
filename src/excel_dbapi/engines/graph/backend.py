"""GraphBackend — WorkbookBackend for Microsoft Graph API."""

from __future__ import annotations

import sys
from typing import Any, cast
from urllib.parse import quote

import httpx

from ...exceptions import NotSupportedError, OperationalError
from ..base import TableData, WorkbookBackend, _normalize_headers
from .auth import TokenProvider, normalize_token_provider
from .client import GraphClient
from .locator import GraphWorkbookLocator, parse_msgraph_dsn
from .session import WorkbookSession


def _col_letter(index: int) -> str:
    """Convert a 0-based column index to an Excel column letter (A, B, ..., Z, AA, AB, ...).

    Args:
        index: 0-based column index.

    Returns:
        Excel-style column letter string.

    Raises:
        ValueError: If *index* is negative.
    """
    if index < 0:
        raise ValueError(f"Column index must be non-negative, got {index}")
    result = ""
    n = index
    while True:
        result = chr(ord("A") + n % 26) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result


def _encode_path_segment(value: str) -> str:
    return quote(value, safe="")


class GraphBackend(WorkbookBackend):
    """Backend that accesses Excel data via Microsoft Graph API.

    ``file_path`` must be a ``msgraph://drives/{drive_id}/items/{item_id}``
    DSN.

    By default the backend is **read-only**.  Pass ``readonly=False`` via
    ``backend_options`` (or the ``connect()`` call) to enable write operations
    (INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE).  Writable mode uses a
    Graph API session with ``persistChanges=true``; changes are applied
    immediately and **cannot be rolled back**.

    Networking options (via ``backend_options`` / ``connect()`` kwargs):
    - ``timeout`` (float, default 30.0): HTTP request timeout in seconds.
    - ``max_retries`` (int, default 3): Number of retries for retryable GETs.
    - ``backoff_factor`` (float, default 0.5): Exponential retry backoff factor.
    """

    supports_transactions: bool = False
    _CONFLICT_STRATEGIES = frozenset({"fail", "force"})
    _WRITE_METHODS = frozenset({"POST", "PATCH", "PUT", "DELETE"})
    _FULL_REWRITE_THRESHOLD = 0.5

    def __init__(
        self,
        file_path: str,
        *,
        data_only: bool = True,
        create: bool = False,
        sanitize_formulas: bool = True,
        credential: Any = None,
        transport: httpx.BaseTransport | None = None,
        readonly: bool = True,
        timeout: float = 30.0,
        max_retries: int = 3,
        backoff_factor: float = 0.5,
        conflict_strategy: str = "fail",
        **options: Any,
    ) -> None:
        if create:
            raise NotSupportedError(
                "Graph backend does not support creating workbooks (create=True)"
            )
        if not data_only:
            raise NotSupportedError(
                "Graph backend does not support formula access (data_only=False)"
            )

        super().__init__(
            file_path,
            data_only=data_only,
            create=create,
            sanitize_formulas=sanitize_formulas,
            **options,
        )

        # Instance attribute — toggleable per connection
        self.readonly: bool = readonly
        if conflict_strategy not in self._CONFLICT_STRATEGIES:
            allowed = ", ".join(sorted(self._CONFLICT_STRATEGIES))
            raise ValueError(
                f"Invalid conflict_strategy {conflict_strategy!r}. "
                f"Expected one of: {allowed}"
            )
        self._conflict_strategy = conflict_strategy
        self._etag: str | None = None

        self._locator: GraphWorkbookLocator = parse_msgraph_dsn(file_path)
        self._token_provider: TokenProvider = normalize_token_provider(credential)
        self._client: GraphClient = GraphClient(
            self._token_provider,
            transport=transport,
            timeout=timeout,
            max_retries=max_retries,
            backoff_factor=backoff_factor,
        )
        self._session: WorkbookSession = WorkbookSession(
            self._client,
            self._locator,
            persist_changes=not readonly,
        )

        # Cache: name → worksheet id
        self._sheet_ids: dict[str, str] = {}
        self._sheets_loaded: bool = False

    # -- WorkbookBackend interface -------------------------------------------

    def load(self) -> None:
        """Fetch worksheet listing (lazy — called on first read)."""
        self._ensure_session()
        self._load_sheets()

    def save(self) -> None:
        """No-op — Graph writes are auto-persisted via session."""

    def snapshot(self) -> Any:
        """Return an opaque marker; no real state to snapshot."""
        return None

    def restore(self, snapshot: Any) -> None:
        """Close current session and clear cached data."""
        self._session.close()
        self._sheets_loaded = False
        self._sheet_ids.clear()

    def list_sheets(self) -> list[str]:
        self._ensure_session()
        self._load_sheets()
        return list(self._sheet_ids.keys())

    def read_sheet(self, sheet_name: str) -> TableData:
        self._ensure_session()
        self._load_sheets()
        ws_id = self._sheet_ids.get(sheet_name)
        if ws_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")

        path = (
            f"{self._locator.item_path}/workbook"
            f"/worksheets/{_encode_path_segment(ws_id)}/usedRange(valuesOnly=true)?$select=values"
        )
        resp = self._session_aware_request("GET", path)
        values = resp.json().get("values", [])

        if not values:
            return TableData(headers=[], rows=[])

        headers = _normalize_headers(values[0])
        rows = [list(row) for row in values[1:]]
        self._check_row_limit(sheet_name, len(rows))
        approx_bytes = sys.getsizeof(headers)
        for row in rows:
            approx_bytes += sys.getsizeof(row)
            approx_bytes += sum(sys.getsizeof(value) for value in row)
        self._check_memory_limit(sheet_name, approx_bytes)
        return TableData(headers=headers, rows=rows)

    # -- Mutating operations -------------------------------------------------

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        """Write table data, using targeted updates/deletes when safe."""
        self._ensure_writable("write_sheet")
        self._ensure_session()
        self._load_sheets()
        ws_id = self._sheet_ids.get(sheet_name)
        if ws_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")

        # Read old used range to know both old row count and column width
        old_values = self._read_used_range(ws_id)
        old_row_count = len(old_values) if old_values else 0
        old_col_count = len(old_values[0]) if old_values else 0

        # Build the full matrix (header + data rows)
        num_cols = len(data.headers)
        matrix: list[list[Any]] = [list(data.headers)]
        for row in data.rows:
            # Pad/trim to header width for rectangular payload
            padded = list(row) + [None] * (num_cols - len(row))
            matrix.append(padded[:num_cols])

        if self._try_patch_changed_rows(ws_id, old_values, matrix, num_cols):
            return

        if self._try_delete_rows(ws_id, old_values, matrix, num_cols):
            return

        self._rewrite_sheet(ws_id, matrix, old_row_count, old_col_count, num_cols)

    def _rewrite_sheet(
        self,
        ws_id: str,
        matrix: list[list[Any]],
        old_row_count: int,
        old_col_count: int,
        num_cols: int,
    ) -> None:
        """Rewrite sheet matrix and clear stale tails/columns."""
        new_row_count = len(matrix)  # header + data rows
        last_col = _col_letter(num_cols - 1) if num_cols > 0 else "A"

        if num_cols > 0 and new_row_count > 0:
            address = f"A1:{last_col}{new_row_count}"
            patch_path = (
                f"{self._locator.item_path}/workbook"
                f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{address}')"
            )
            self._session_aware_request("PATCH", patch_path, json={"values": matrix})

        max_col_count = max(old_col_count, num_cols) if old_col_count else num_cols
        tail_last_col = _col_letter(max_col_count - 1) if max_col_count > 0 else "A"
        if old_row_count > new_row_count and max_col_count > 0:
            tail_start = new_row_count + 1
            tail_address = f"A{tail_start}:{tail_last_col}{old_row_count}"
            clear_path = (
                f"{self._locator.item_path}/workbook"
                f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{tail_address}')/clear"
            )
            self._session_aware_request(
                "POST", clear_path, json={"applyTo": "Contents"}
            )

        if old_col_count > num_cols and new_row_count > 0:
            right_start_col = _col_letter(num_cols)
            right_end_col = _col_letter(old_col_count - 1)
            right_address = f"{right_start_col}1:{right_end_col}{new_row_count}"
            clear_right_path = (
                f"{self._locator.item_path}/workbook"
                f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{right_address}')/clear"
            )
            self._session_aware_request(
                "POST", clear_right_path, json={"applyTo": "Contents"}
            )

    def _try_patch_changed_rows(
        self,
        ws_id: str,
        old_values: list[list[Any]],
        matrix: list[list[Any]],
        num_cols: int,
    ) -> bool:
        """Patch only changed rows when old/new shapes match."""
        if not old_values or num_cols == 0:
            return False
        if len(old_values) != len(matrix):
            return False

        old_headers = list(old_values[0]) if old_values else []
        if old_headers != matrix[0]:
            return False

        changed_rows: list[int] = []
        for idx, (old_row, new_row) in enumerate(
            zip(old_values[1:], matrix[1:]), start=2
        ):
            old_rect = self._rect_row(old_row, num_cols)
            if old_rect != new_row:
                changed_rows.append(idx)

        if not changed_rows:
            return True

        total_data_rows = len(matrix) - 1
        if total_data_rows <= 0:
            return False
        if (len(changed_rows) / total_data_rows) > self._FULL_REWRITE_THRESHOLD:
            return False

        row_groups = self._group_consecutive(changed_rows)
        last_col = _col_letter(num_cols - 1)
        for start_row, end_row in row_groups:
            values = [
                matrix[row_number - 1] for row_number in range(start_row, end_row + 1)
            ]
            address = f"A{start_row}:{last_col}{end_row}"
            patch_path = (
                f"{self._locator.item_path}/workbook"
                f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{address}')"
            )
            self._session_aware_request("PATCH", patch_path, json={"values": values})
        return True

    def _try_delete_rows(
        self,
        ws_id: str,
        old_values: list[list[Any]],
        matrix: list[list[Any]],
        num_cols: int,
    ) -> bool:
        """Delete removed rows with Graph range delete when possible."""
        if not old_values or num_cols == 0:
            return False
        if len(matrix) >= len(old_values):
            return False

        old_headers = list(old_values[0]) if old_values else []
        if old_headers != matrix[0]:
            return False

        old_rows = [self._rect_row(row, num_cols) for row in old_values[1:]]
        new_rows = matrix[1:]
        deleted_data_rows = self._find_deleted_row_indices(old_rows, new_rows)
        if not deleted_data_rows:
            return False

        deleted_sheet_rows = [idx + 2 for idx in deleted_data_rows]
        row_groups = self._group_consecutive(deleted_sheet_rows)
        last_col = _col_letter(num_cols - 1)

        for start_row, end_row in reversed(row_groups):
            address = f"A{start_row}:{last_col}{end_row}"
            delete_path = (
                f"{self._locator.item_path}/workbook"
                f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{address}')/delete"
            )
            self._session_aware_request("POST", delete_path, json={"shift": "Up"})
        return True

    @staticmethod
    def _rect_row(row: list[Any], width: int) -> list[Any]:
        padded = list(row) + [None] * (width - len(row))
        return padded[:width]

    @staticmethod
    def _group_consecutive(rows: list[int]) -> list[tuple[int, int]]:
        if not rows:
            return []
        groups: list[tuple[int, int]] = []
        start = rows[0]
        end = rows[0]
        for row in rows[1:]:
            if row == end + 1:
                end = row
                continue
            groups.append((start, end))
            start = row
            end = row
        groups.append((start, end))
        return groups

    @staticmethod
    def _find_deleted_row_indices(
        old_rows: list[list[Any]], new_rows: list[list[Any]]
    ) -> list[int]:
        """Return old-row indexes removed from old_rows to form new_rows."""
        deleted: list[int] = []
        old_idx = 0
        new_idx = 0
        while old_idx < len(old_rows) and new_idx < len(new_rows):
            if old_rows[old_idx] == new_rows[new_idx]:
                old_idx += 1
                new_idx += 1
                continue
            deleted.append(old_idx)
            old_idx += 1
        if new_idx != len(new_rows):
            return []
        while old_idx < len(old_rows):
            deleted.append(old_idx)
            old_idx += 1
        return deleted

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        """Append a single row to *sheet_name* and return the 1-based row index."""
        self._ensure_writable("append_row")
        self._ensure_session()
        self._load_sheets()
        ws_id = self._sheet_ids.get(sheet_name)
        if ws_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")

        # Single read to get both row count and header width
        values = self._read_used_range(ws_id)
        if not values:
            # Empty sheet — write to row 1
            next_row = 1
            num_cols = len(row)
        else:
            next_row = len(values) + 1  # 1-based; usedRange includes header
            num_cols = len(values[0])

        last_col = _col_letter(num_cols - 1) if num_cols > 0 else "A"

        # Pad/trim row to header width
        padded = list(row) + [None] * (num_cols - len(row))
        row_values = [padded[:num_cols]]

        address = f"A{next_row}:{last_col}{next_row}"
        patch_path = (
            f"{self._locator.item_path}/workbook"
            f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{address}')"
        )
        self._session_aware_request("PATCH", patch_path, json={"values": row_values})
        return next_row

    def create_sheet(self, name: str, headers: list[str]) -> None:
        """Create a new worksheet and write the header row."""
        self._ensure_writable("create_sheet")
        self._ensure_session()

        # POST to add worksheet
        ws_path = f"{self._locator.item_path}/workbook/worksheets/add"
        resp = self._session_aware_request("POST", ws_path, json={"name": name})
        ws_info = resp.json()
        ws_id = ws_info["id"]

        # Write header row
        if headers:
            num_cols = len(headers)
            last_col = _col_letter(num_cols - 1)
            address = f"A1:{last_col}1"
            header_path = (
                f"{self._locator.item_path}/workbook"
                f"/worksheets/{_encode_path_segment(ws_id)}/range(address='{address}')"
            )
            self._session_aware_request(
                "PATCH", header_path, json={"values": [headers]}
            )

        # Invalidate cache — next _load_sheets() will re-fetch all sheets
        self._invalidate_sheet_cache()

    def drop_sheet(self, name: str) -> None:
        """Delete a worksheet by name."""
        self._ensure_writable("drop_sheet")
        self._ensure_session()
        self._load_sheets()
        ws_id = self._sheet_ids.get(name)
        if ws_id is None:
            raise ValueError(f"Sheet '{name}' not found in Excel")

        delete_path = f"{self._locator.item_path}/workbook/worksheets/{_encode_path_segment(ws_id)}"
        self._session_aware_request("DELETE", delete_path)

        # Invalidate cache
        self._invalidate_sheet_cache()

    # -- Internal helpers ----------------------------------------------------

    def _ensure_writable(self, operation: str) -> None:
        """Raise NotSupportedError if backend is read-only."""
        if self.readonly:
            raise NotSupportedError(
                f"{operation} is not supported by the read-only Graph backend"
            )

    def _ensure_session(self) -> None:
        was_open = self._session.is_open
        self._session.ensure_open()
        if not was_open:
            self._prime_workbook_etag()

    def _session_aware_request(
        self, method: str, path: str, **kwargs: Any
    ) -> httpx.Response:
        """HTTP request with stale-session recovery.

        If the request fails with a session-related error, close and reopen
        the session, invalidate cached sheet data, and retry once.

        Stale-session recovery is attempted for all HTTP methods.  The retry
        targets the *session infrastructure* (expired session ID), not
        transient server errors — so it is safe even for mutating methods.
        """
        method_upper = method.upper()
        dispatch = {
            "GET": self._client.get,
            "POST": self._client.post,
            "PATCH": self._client.patch,
            "DELETE": self._client.delete,
        }
        send = dispatch[method_upper]
        headers = dict(cast(dict[str, str], kwargs.pop("headers", {})))
        if (
            self._conflict_strategy == "fail"
            and method_upper in self._WRITE_METHODS
            and self._etag is not None
        ):
            headers["If-Match"] = self._etag
        if headers:
            kwargs["headers"] = headers
        try:
            response = send(path, **kwargs)
            self._update_etag_from_response(response)
            return response
        except OperationalError as exc:
            if (
                self._conflict_strategy == "fail"
                and method_upper in self._WRITE_METHODS
                and self._is_conflict_error(exc)
            ):
                raise OperationalError(
                    "Concurrent modification detected: workbook was modified by another session"
                ) from exc
            if not self._is_session_error(exc):
                raise
        # Session expired — reopen and retry once
        self._session.reopen()
        self._sheets_loaded = False
        self._sheet_ids.clear()
        self._load_sheets()
        response = send(path, **kwargs)
        self._update_etag_from_response(response)
        return response

    def _prime_workbook_etag(self) -> None:
        if self._conflict_strategy != "fail":
            return
        path = f"{self._locator.item_path}/workbook"
        try:
            self._session_aware_request("GET", path)
        except OperationalError:
            return

    def _update_etag_from_response(self, response: httpx.Response) -> None:
        etag = response.headers.get("ETag")
        if etag:
            self._etag = etag
            return
        try:
            payload = response.json()
        except ValueError:
            return
        if isinstance(payload, dict):
            odata_etag = payload.get("@odata.etag")
            if isinstance(odata_etag, str) and odata_etag:
                self._etag = odata_etag

    @staticmethod
    def _is_conflict_error(exc: OperationalError) -> bool:
        msg = str(exc)
        return "412" in msg or "precondition failed" in msg.lower()

    @staticmethod
    def _is_session_error(exc: OperationalError) -> bool:
        """Check if the error indicates an expired/invalid session."""
        msg = str(exc)
        msg_lower = msg.lower()
        # Match Graph-specific invalid session indicators
        return (
            "invalidSessionId" in msg
            or "invalidsession" in msg_lower
            or ("404" in msg_lower and "session" in msg_lower)
        )

    def _load_sheets(self) -> None:
        if self._sheets_loaded:
            return
        path = f"{self._locator.item_path}/workbook/worksheets?$select=id,name"
        resp = self._session_aware_request("GET", path)
        self._sheet_ids.clear()
        for ws in resp.json()["value"]:
            self._sheet_ids[ws["name"]] = ws["id"]
        self._sheets_loaded = True

    def _invalidate_sheet_cache(self) -> None:
        """Clear cached worksheet list so next access re-fetches."""
        self._sheets_loaded = False
        self._sheet_ids.clear()

    def _used_range_row_count(self, ws_id: str) -> int:
        """Return the total number of used rows (including header) for a worksheet."""
        values = self._read_used_range(ws_id)
        return len(values) if values else 0

    def _read_used_range(self, ws_id: str) -> list[list[Any]]:
        """Return raw values matrix from usedRange, or empty list."""
        path = (
            f"{self._locator.item_path}/workbook"
            f"/worksheets/{_encode_path_segment(ws_id)}/usedRange(valuesOnly=true)?$select=values"
        )
        resp = self._session_aware_request("GET", path)
        return cast(list[list[Any]], resp.json().get("values", []))

    def close(self) -> None:
        """Close session and HTTP client."""
        self._session.close()
        self._client.close()
