"""Tests for GraphBackend — read, write, create, drop, session, cache."""

import json
from typing import Any

import httpx
import pytest

from excel_dbapi.engines.graph.backend import GraphBackend, _col_letter
from excel_dbapi.exceptions import NotSupportedError


DSN = "msgraph://drives/drv-1/items/itm-1"


# ---------------------------------------------------------------------------
# Helpers — mock transport builders
# ---------------------------------------------------------------------------


def _graph_handler(request: httpx.Request) -> httpx.Response:
    """Mock handler covering session + worksheets + usedRange."""
    path = request.url.path

    if path.endswith("/createSession"):
        return httpx.Response(201, json={"id": "sess-mock"})

    if path.endswith("/closeSession"):
        return httpx.Response(204)

    if path.endswith("/worksheets") or "/worksheets?" in str(request.url):
        return httpx.Response(
            200,
            json={
                "value": [
                    {"id": "ws-1", "name": "Users"},
                    {"id": "ws-2", "name": "Orders"},
                ]
            },
        )

    if "usedRange" in path:
        # Determine which sheet was requested
        if "ws-1" in path:
            return httpx.Response(
                200,
                json={
                    "values": [
                        ["id", "name", "email"],
                        [1, "Ada", "ada@example.com"],
                        [2, "Bob", "bob@example.com"],
                    ]
                },
            )
        if "ws-2" in path:
            return httpx.Response(
                200,
                json={
                    "values": [
                        ["order_id", "user_id", "total"],
                        [101, 1, 99.99],
                    ]
                },
            )

    return httpx.Response(404)


def _make_backend(**kwargs: Any) -> GraphBackend:
    transport = httpx.MockTransport(_graph_handler)
    return GraphBackend(
        DSN,
        credential="test-token",
        transport=transport,
        **kwargs,
    )


def _writable_handler_factory():
    """Build a stateful mock handler that records writes and supports full CRUD."""
    state: dict[str, Any] = {
        "worksheets": {
            "ws-1": {
                "name": "Users",
                "values": [
                    ["id", "name", "email"],
                    [1, "Ada", "ada@example.com"],
                    [2, "Bob", "bob@example.com"],
                ],
            },
        },
        "requests": [],  # record of (method, path, body)
    }

    def handler(request: httpx.Request) -> httpx.Response:
        method = request.method
        path = request.url.path
        body = None
        if request.content:
            try:
                body = json.loads(request.content)
            except (json.JSONDecodeError, UnicodeDecodeError):
                body = None

        state["requests"].append((method, path, body))

        # Session management
        if path.endswith("/createSession"):
            return httpx.Response(201, json={"id": "sess-rw"})
        if path.endswith("/closeSession"):
            return httpx.Response(204)

        # Worksheet listing
        if (
            path.endswith("/worksheets") or "/worksheets?" in str(request.url)
        ) and method == "GET":
            ws_list = [
                {"id": ws_id, "name": ws_data["name"]}
                for ws_id, ws_data in state["worksheets"].items()
            ]
            return httpx.Response(200, json={"value": ws_list})

        # Add worksheet
        if path.endswith("/worksheets/add") and method == "POST":
            new_name = body["name"] if body else "NewSheet"
            new_id = f"ws-new-{len(state['worksheets']) + 1}"
            state["worksheets"][new_id] = {"name": new_name, "values": []}
            return httpx.Response(201, json={"id": new_id, "name": new_name})

        # Delete worksheet
        for ws_id in list(state["worksheets"]):
            if path.endswith(f"/worksheets/{ws_id}") and method == "DELETE":
                del state["worksheets"][ws_id]
                return httpx.Response(204)

        # usedRange
        if "usedRange" in path and method == "GET":
            for ws_id, ws_data in state["worksheets"].items():
                if ws_id in path:
                    return httpx.Response(200, json={"values": ws_data["values"]})
            return httpx.Response(200, json={"values": []})

        # PATCH range (write data)
        if "/range(" in path and method == "PATCH":
            return httpx.Response(200, json=body or {})

        if path.endswith("/delete") and method == "POST":
            return httpx.Response(200, json={})

        # POST range/clear
        if path.endswith("/clear") and method == "POST":
            return httpx.Response(200, json={})

        return httpx.Response(404)

    return handler, state


def _make_writable_backend(**kwargs: Any):
    """Create a writable backend + its state tracker."""
    handler, state = _writable_handler_factory()
    transport = httpx.MockTransport(handler)
    backend = GraphBackend(
        DSN,
        credential="test-token",
        transport=transport,
        readonly=False,
        **kwargs,
    )
    return backend, state


# ---------------------------------------------------------------------------
# Unit test: _col_letter helper
# ---------------------------------------------------------------------------


class TestColLetter:
    def test_single_letters(self):
        assert _col_letter(0) == "A"
        assert _col_letter(1) == "B"
        assert _col_letter(25) == "Z"

    def test_double_letters(self):
        assert _col_letter(26) == "AA"
        assert _col_letter(27) == "AB"
        assert _col_letter(51) == "AZ"
        assert _col_letter(52) == "BA"

    def test_triple_letters(self):
        assert _col_letter(702) == "AAA"


# ---------------------------------------------------------------------------
# Original read-only tests (preserved from v1.2)
# ---------------------------------------------------------------------------


class TestGraphBackendInit:
    def test_create_rejected(self):
        with pytest.raises(NotSupportedError, match="create"):
            _make_backend(create=True)

    def test_data_only_false_rejected(self):
        with pytest.raises(NotSupportedError, match="data_only"):
            _make_backend(data_only=False)

    def test_readonly_flag_default(self):
        backend = _make_backend()
        assert backend.readonly is True

    def test_readonly_false(self):
        backend, _ = _make_writable_backend()
        assert backend.readonly is False
        backend.close()

    def test_supports_transactions_false(self):
        backend = _make_backend()
        assert backend.supports_transactions is False


class TestGraphBackendListSheets:
    def test_list_sheets(self):
        backend = _make_backend()
        sheets = backend.list_sheets()
        assert "Users" in sheets
        assert "Orders" in sheets
        assert len(sheets) == 2


class TestGraphBackendReadSheet:
    def test_read_users(self):
        backend = _make_backend()
        data = backend.read_sheet("Users")
        assert data.headers == ["id", "name", "email"]
        assert len(data.rows) == 2
        assert data.rows[0] == [1, "Ada", "ada@example.com"]
        assert data.rows[1] == [2, "Bob", "bob@example.com"]

    def test_read_orders(self):
        backend = _make_backend()
        data = backend.read_sheet("Orders")
        assert data.headers == ["order_id", "user_id", "total"]
        assert len(data.rows) == 1
        assert data.rows[0] == [101, 1, 99.99]

    def test_read_unknown_sheet(self):
        backend = _make_backend()
        with pytest.raises(ValueError, match="not found"):
            backend.read_sheet("Missing")

    def test_read_sheet_enforces_row_limit(self):
        backend = _make_backend(max_rows=1)
        with pytest.raises(Exception, match="max_rows"):
            backend.read_sheet("Users")

    def test_read_sheet_enforces_memory_limit(self):
        backend = _make_backend(max_memory_mb=0.00001)
        with pytest.raises(Exception, match="max_memory_mb"):
            backend.read_sheet("Users")


class TestGraphBackendReadOnly:
    def test_write_sheet_raises(self):
        backend = _make_backend()
        backend.load()
        with pytest.raises(NotSupportedError, match="read-only"):
            backend.write_sheet("Users", backend.read_sheet("Users"))

    def test_append_row_raises(self):
        backend = _make_backend()
        backend.load()
        with pytest.raises(NotSupportedError, match="read-only"):
            backend.append_row("Users", [3, "Carol", "carol@example.com"])

    def test_create_sheet_raises(self):
        backend = _make_backend()
        backend.load()
        with pytest.raises(NotSupportedError, match="read-only"):
            backend.create_sheet("NewSheet", ["a", "b"])

    def test_drop_sheet_raises(self):
        backend = _make_backend()
        backend.load()
        with pytest.raises(NotSupportedError, match="read-only"):
            backend.drop_sheet("Users")


class TestGraphBackendSnapshot:
    def test_snapshot_returns_none(self):
        backend = _make_backend()
        assert backend.snapshot() is None

    def test_restore_clears_cache(self):
        backend = _make_backend()
        backend.list_sheets()  # populate cache
        assert backend._sheets_loaded is True
        backend.restore(None)
        assert backend._sheets_loaded is False


class TestGraphBackendSave:
    def test_save_is_noop(self):
        backend = _make_backend()
        backend.save()  # should not raise


class TestGraphBackendClose:
    def test_close_releases_resources(self):
        """close() should close both session and HTTP client."""
        backend = _make_backend()
        backend.load()  # opens session
        assert backend._session.is_open is True
        backend.close()
        assert backend._session.is_open is False

    def test_double_close_is_safe(self):
        """Calling close() twice should not raise."""
        backend = _make_backend()
        backend.load()
        backend.close()
        backend.close()  # second close is a no-op


class TestGraphBackendSessionRecovery:
    def test_session_expired_reopens(self):
        """If a read fails with a session error, backend should reopen."""
        call_count = {"n": 0}

        def handler(request: httpx.Request) -> httpx.Response:
            path = request.url.path
            if path.endswith("/createSession"):
                call_count["n"] += 1
                return httpx.Response(201, json={"id": f"sess-{call_count['n']}"})
            if path.endswith("/closeSession"):
                return httpx.Response(204)
            if path.endswith("/worksheets") or "/worksheets?" in str(request.url):
                return httpx.Response(
                    200, json={"value": [{"id": "ws-1", "name": "Data"}]}
                )
            if "usedRange" in path:
                # First usedRange call fails with 404 (session expired)
                if call_count.get("read", 0) == 0:
                    call_count["read"] = 1
                    return httpx.Response(
                        404, json={"error": {"code": "invalidSessionId"}}
                    )
                return httpx.Response(200, json={"values": [["a", "b"], [1, 2]]})
            return httpx.Response(404)

        transport = httpx.MockTransport(handler)
        backend = GraphBackend(DSN, credential="tok", transport=transport)
        data = backend.read_sheet("Data")
        assert data.headers == ["a", "b"]
        # Session was opened twice (original + recovery)
        assert call_count["n"] == 2
        backend.close()


# ---------------------------------------------------------------------------
# v1.3 Write operation tests
# ---------------------------------------------------------------------------


class TestGraphBackendWriteSession:
    def test_writable_backend_uses_persist_changes(self):
        """Writable backend should create session with persistChanges=true."""
        backend, state = _make_writable_backend()
        backend.load()  # triggers session creation
        # Find the createSession request
        create_reqs = [r for r in state["requests"] if r[1].endswith("/createSession")]
        assert len(create_reqs) >= 1
        assert create_reqs[0][2] == {"persistChanges": True}
        backend.close()

    def test_readonly_backend_uses_no_persist(self):
        """Read-only backend should create session with persistChanges=false."""
        handler, state = _writable_handler_factory()
        transport = httpx.MockTransport(handler)
        backend = GraphBackend(
            DSN, credential="tok", transport=transport, readonly=True
        )
        backend.load()
        create_reqs = [r for r in state["requests"] if r[1].endswith("/createSession")]
        assert len(create_reqs) >= 1
        assert create_reqs[0][2] == {"persistChanges": False}
        backend.close()


class TestGraphBackendAppendRow:
    def test_append_row_basic(self):
        backend, state = _make_writable_backend()
        row_idx = backend.append_row("Users", [3, "Carol", "carol@example.com"])
        # usedRange has 3 rows (1 header + 2 data), so next row = 4
        assert row_idx == 4
        # Verify PATCH was sent
        patch_reqs = [r for r in state["requests"] if r[0] == "PATCH"]
        assert len(patch_reqs) >= 1
        last_patch = patch_reqs[-1]
        assert "range(address=" in last_patch[1]
        assert last_patch[2]["values"] == [[3, "Carol", "carol@example.com"]]
        backend.close()

    def test_append_row_unknown_sheet(self):
        backend, _ = _make_writable_backend()
        with pytest.raises(ValueError, match="not found"):
            backend.append_row("NoSuchSheet", [1, 2, 3])
        backend.close()

    def test_append_row_pads_short_row(self):
        """Row shorter than header width is padded with None."""
        backend, state = _make_writable_backend()
        backend.append_row("Users", [3, "Carol"])
        patch_reqs = [r for r in state["requests"] if r[0] == "PATCH"]
        last_patch = patch_reqs[-1]
        # Should be padded to 3 columns
        assert last_patch[2]["values"] == [[3, "Carol", None]]
        backend.close()


class TestGraphBackendWriteSheet:
    def test_write_sheet_full_overwrite(self):
        from excel_dbapi.engines.base import TableData

        backend, state = _make_writable_backend()
        new_data = TableData(
            headers=["id", "name", "email"],
            rows=[
                [1, "Ada", "ada-new@example.com"],
            ],
        )
        backend.write_sheet("Users", new_data)

        # Should have PATCH (overwrite) + POST (clear tail)
        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        assert len(patch_reqs) >= 1
        # The PATCH payload should have header + 1 data row
        patch_body = patch_reqs[-1][2]
        assert len(patch_body["values"]) == 2  # header + 1 row

        # Old had 3 rows, new has 2 → tail clear should happen
        clear_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/clear" in r[1]
        ]
        assert len(clear_reqs) == 1
        assert clear_reqs[0][2] == {"applyTo": "Contents"}
        backend.close()

    def test_write_sheet_no_tail_clear_when_same_size(self):
        """No clear needed when new data has same or more rows."""
        from excel_dbapi.engines.base import TableData

        backend, state = _make_writable_backend()
        new_data = TableData(
            headers=["id", "name", "email"],
            rows=[
                [1, "Ada", "ada@example.com"],
                [2, "Bob", "bob@example.com"],
                [3, "Carol", "carol@example.com"],
            ],
        )
        backend.write_sheet("Users", new_data)

        clear_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/clear" in r[1]
        ]
        assert len(clear_reqs) == 0
        backend.close()

    def test_write_sheet_unknown_sheet(self):
        from excel_dbapi.engines.base import TableData

        backend, _ = _make_writable_backend()
        with pytest.raises(ValueError, match="not found"):
            backend.write_sheet("NoSuchSheet", TableData(headers=["a"], rows=[]))
        backend.close()

    def test_write_sheet_pads_rows(self):
        """Rows shorter than header width are padded to ensure rectangular payload."""
        from excel_dbapi.engines.base import TableData

        backend, state = _make_writable_backend()
        new_data = TableData(
            headers=["id", "name", "email"],
            rows=[
                [1, "Ada"],  # short row — missing email
            ],
        )
        backend.write_sheet("Users", new_data)

        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        values = patch_reqs[-1][2]["values"]
        # Header row + 1 data row, data row padded to 3 cols
        assert values[1] == [1, "Ada", None]
        backend.close()


class TestGraphBackendCreateSheet:
    def test_create_sheet_basic(self):
        backend, state = _make_writable_backend()
        backend.create_sheet("Products", ["sku", "name", "price"])

        # Should have POST to /worksheets/add
        add_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/worksheets/add" in r[1]
        ]
        assert len(add_reqs) == 1
        assert add_reqs[0][2]["name"] == "Products"

        # Should have PATCH for header row
        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        assert len(patch_reqs) >= 1
        last_patch = patch_reqs[-1]
        assert last_patch[2]["values"] == [["sku", "name", "price"]]

        backend.close()

    def test_create_sheet_empty_headers(self):
        """Creating a sheet with no headers should still create the worksheet."""
        backend, state = _make_writable_backend()
        backend.create_sheet("Empty", [])

        add_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/worksheets/add" in r[1]
        ]
        assert len(add_reqs) == 1

        # No PATCH for empty headers
        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        assert len(patch_reqs) == 0
        backend.close()

    def test_create_sheet_invalidates_cache(self):
        backend, _ = _make_writable_backend()
        backend.list_sheets()  # populate cache
        assert backend._sheets_loaded is True
        backend.create_sheet("NewSheet", ["a", "b"])
        # After create, cache should be fully invalidated
        assert backend._sheets_loaded is False
        assert len(backend._sheet_ids) == 0
        backend.close()

    def test_create_then_list_includes_old_and_new_sheets(self):
        """After CREATE, list_sheets() re-fetches and returns all sheets."""
        backend, _ = _make_writable_backend()
        old_sheets = backend.list_sheets()
        assert "Users" in old_sheets
        backend.create_sheet("Products", ["sku", "name"])
        # Cache was invalidated; list_sheets re-fetches from server
        all_sheets = backend.list_sheets()
        assert "Users" in all_sheets  # old sheet still visible
        assert "Products" in all_sheets  # new sheet also visible
        backend.close()

    def test_create_then_read_existing_sheet(self):
        """After CREATE, reading an existing sheet still works."""
        backend, _ = _make_writable_backend()
        backend.create_sheet("Products", ["sku", "name"])
        # Should be able to read the original Users sheet
        data = backend.read_sheet("Users")
        assert data.headers == ["id", "name", "email"]
        assert len(data.rows) == 2
        backend.close()


class TestGraphBackendDropSheet:
    def test_drop_sheet_basic(self):
        backend, state = _make_writable_backend()
        backend.list_sheets()  # populate cache
        backend.drop_sheet("Users")

        delete_reqs = [r for r in state["requests"] if r[0] == "DELETE"]
        assert len(delete_reqs) == 1
        assert "ws-1" in delete_reqs[0][1]
        backend.close()

    def test_drop_sheet_unknown(self):
        backend, _ = _make_writable_backend()
        with pytest.raises(ValueError, match="not found"):
            backend.drop_sheet("NoSuchSheet")
        backend.close()

    def test_drop_sheet_invalidates_cache(self):
        backend, _ = _make_writable_backend()
        backend.list_sheets()
        assert "Users" in backend._sheet_ids
        backend.drop_sheet("Users")
        assert backend._sheets_loaded is False
        backend.close()


class TestGraphBackendSessionRecoveryWrite:
    """Session recovery on write (PATCH/POST/DELETE) requests."""

    def test_session_expired_on_patch_reopens(self):
        """Stale session error on PATCH triggers session reopen."""
        call_count = {"sessions": 0, "patches": 0}

        def handler(request: httpx.Request) -> httpx.Response:
            path = request.url.path
            method = request.method

            if path.endswith("/createSession"):
                call_count["sessions"] += 1
                return httpx.Response(
                    201, json={"id": f"sess-{call_count['sessions']}"}
                )
            if path.endswith("/closeSession"):
                return httpx.Response(204)
            if path.endswith("/worksheets") or "/worksheets?" in str(request.url):
                return httpx.Response(
                    200, json={"value": [{"id": "ws-1", "name": "Data"}]}
                )
            if "usedRange" in path:
                return httpx.Response(200, json={"values": [["a", "b"], [1, 2]]})
            if "/range(" in path and method == "PATCH":
                call_count["patches"] += 1
                if call_count["patches"] == 1:
                    # First PATCH fails with session error
                    return httpx.Response(
                        404, json={"error": {"code": "InvalidSessionId"}}
                    )
                return httpx.Response(200, json={})
            return httpx.Response(404)

        transport = httpx.MockTransport(handler)
        backend = GraphBackend(
            DSN, credential="tok", transport=transport, readonly=False
        )
        # append_row triggers a PATCH
        backend.append_row("Data", [3, 4])
        # Session was opened twice
        assert call_count["sessions"] == 2
        backend.close()


class TestColLetterNegative:
    def test_negative_index_raises(self):
        with pytest.raises(ValueError, match="non-negative"):
            _col_letter(-1)

    def test_negative_large_raises(self):
        with pytest.raises(ValueError, match="non-negative"):
            _col_letter(-100)


class TestGraphBackendWriteSheetColumnNarrowing:
    """write_sheet clears stale right-side columns when new data is narrower."""

    def test_write_fewer_columns_clears_right_side(self):
        from excel_dbapi.engines.base import TableData

        backend, state = _make_writable_backend()
        # Original Users sheet has 3 columns: id, name, email
        # Write with only 2 columns
        new_data = TableData(
            headers=["id", "name"],
            rows=[
                [1, "Ada"],
                [2, "Bob"],
            ],
        )
        backend.write_sheet("Users", new_data)

        # Should have a PATCH for the 2-column data
        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        assert len(patch_reqs) >= 1
        # Values should be 2 columns wide
        values = patch_reqs[-1][2]["values"]
        assert len(values[0]) == 2  # header row: ["id", "name"]

        # Should have a POST /clear for right-side stale column (column C)
        clear_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/clear" in r[1]
        ]
        # At least one clear for right-side columns
        right_clear = [r for r in clear_reqs if "C1:" in r[1] or ":C" in r[1]]
        assert len(right_clear) >= 1
        backend.close()

    def test_write_same_columns_no_right_clear(self):
        from excel_dbapi.engines.base import TableData

        backend, state = _make_writable_backend()
        # Same column count as original (3)
        new_data = TableData(
            headers=["id", "name", "email"],
            rows=[
                [1, "Ada", "ada@example.com"],
                [2, "Bob", "bob@example.com"],
            ],
        )
        backend.write_sheet("Users", new_data)

        # No clear should happen — same row count and same column count
        clear_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/clear" in r[1]
        ]
        assert len(clear_reqs) == 0
        backend.close()


class TestGraphBackendLoadSheetsSessionRecovery:
    """_load_sheets now uses _session_aware_request for stale session recovery."""

    def test_stale_session_on_worksheet_listing_recovers(self):
        call_count = {"sessions": 0, "listings": 0}

        def handler(request: httpx.Request) -> httpx.Response:
            path = request.url.path
            if path.endswith("/createSession"):
                call_count["sessions"] += 1
                return httpx.Response(
                    201, json={"id": f"sess-{call_count['sessions']}"}
                )
            if path.endswith("/closeSession"):
                return httpx.Response(204)
            if "/worksheets?" in str(request.url) or path.endswith("/worksheets"):
                call_count["listings"] += 1
                if call_count["listings"] == 1:
                    # First listing fails with session error
                    return httpx.Response(
                        404, json={"error": {"code": "invalidSessionId"}}
                    )
                return httpx.Response(
                    200, json={"value": [{"id": "ws-1", "name": "Data"}]}
                )
            if "usedRange" in path:
                return httpx.Response(200, json={"values": [["a", "b"], [1, 2]]})
            return httpx.Response(404)

        transport = httpx.MockTransport(handler)
        backend = GraphBackend(DSN, credential="tok", transport=transport)
        sheets = backend.list_sheets()
        assert "Data" in sheets
        # Session was opened twice (original + recovery)
        assert call_count["sessions"] == 2
        backend.close()


class TestIsSessionError:
    """_is_session_error matches specific Graph invalid session signals."""

    def test_invalid_session_id(self):
        from excel_dbapi.exceptions import OperationalError

        exc = OperationalError("Graph API error: invalidSessionId")
        assert GraphBackend._is_session_error(exc) is True

    def test_generic_404_not_session(self):
        from excel_dbapi.exceptions import OperationalError

        exc = OperationalError("404: resource not found")
        # Generic 404 without 'session' should NOT trigger recovery
        assert GraphBackend._is_session_error(exc) is False

    def test_404_with_session(self):
        from excel_dbapi.exceptions import OperationalError

        exc = OperationalError("404: session expired")
        assert GraphBackend._is_session_error(exc) is True

    def test_unrelated_error(self):
        from excel_dbapi.exceptions import OperationalError

        exc = OperationalError("timeout connecting to server")
        assert GraphBackend._is_session_error(exc) is False
