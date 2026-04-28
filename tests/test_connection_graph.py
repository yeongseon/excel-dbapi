"""Integration tests: ExcelConnection + Graph backend via mock transport."""
from pathlib import Path
from typing import Any, cast
import json
import sys
import types

from openpyxl import Workbook
import httpx
import pytest

from excel_dbapi import connect
from excel_dbapi.connection import ExcelConnection
from excel_dbapi.engines.graph.auth import _has_get_token_with_args, normalize_token_provider
from excel_dbapi.engines.graph.client import GraphClient
from excel_dbapi.engines.graph.session import WorkbookSession
from excel_dbapi.exceptions import BackendOperationError, DatabaseError, Error, NotSupportedError, OperationalError


DSN = "msgraph://drives/drv-test/items/itm-test"


# ---------------------------------------------------------------------------
# Mock transport — stateful handler for full CRUD
# ---------------------------------------------------------------------------


def _build_handler():
    """Stateful mock handler supporting reads and writes."""
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
            },
        },
        "requests": [],
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

        if path.endswith("/createSession"):
            return httpx.Response(201, json={"id": "sess-conn"})
        if path.endswith("/closeSession"):
            return httpx.Response(204)

        if (
            path.endswith("/worksheets") or "/worksheets?" in str(request.url)
        ) and method == "GET":
            ws_list = [
                {"id": ws_id, "name": ws_data["name"]}
                for ws_id, ws_data in state["worksheets"].items()
            ]
            return httpx.Response(200, json={"value": ws_list})

        if path.endswith("/worksheets/add") and method == "POST":
            new_name = body["name"] if body else "NewSheet"
            new_id = f"ws-new-{len(state['worksheets']) + 1}"
            state["worksheets"][new_id] = {"name": new_name, "values": []}
            return httpx.Response(201, json={"id": new_id, "name": new_name})

        for ws_id in list(state["worksheets"]):
            if path.endswith(f"/worksheets/{ws_id}") and method == "DELETE":
                del state["worksheets"][ws_id]
                return httpx.Response(204)

        if "usedRange" in path and method == "GET":
            for ws_id, ws_data in state["worksheets"].items():
                if ws_id in path:
                    return httpx.Response(200, json={"values": ws_data["values"]})
            return httpx.Response(200, json={"values": []})

        if "/range(" in path and method == "PATCH":
            return httpx.Response(200, json=body or {})

        if path.endswith("/delete") and method == "POST":
            return httpx.Response(200, json={})

        if path.endswith("/clear") and method == "POST":
            return httpx.Response(200, json={})

        return httpx.Response(404)

    return handler, state


def _graph_handler(request: httpx.Request) -> httpx.Response:
    """Simple stateless handler for read-only tests."""
    handler, _ = _build_handler()
    return handler(request)


def _make_connection(**kwargs) -> ExcelConnection:
    transport = httpx.MockTransport(_graph_handler)
    kwargs.setdefault("engine", None)  # auto-detect from DSN
    return ExcelConnection(
        DSN,
        credential="test-token",
        transport=transport,
        **kwargs,
    )


def _make_rw_connection(**kwargs):
    """Create a writable connection + state tracker."""
    handler, state = _build_handler()
    transport = httpx.MockTransport(handler)
    kwargs.setdefault("engine", None)
    conn = ExcelConnection(
        DSN,
        credential="test-token",
        transport=transport,
        readonly=False,
        **kwargs,
    )
    return conn, state


# ---------------------------------------------------------------------------
# Read-only tests (preserved from v1.2)
# ---------------------------------------------------------------------------


class TestGraphConnectionAutoDetect:
    def test_engine_auto_detected_from_dsn(self):
        conn = _make_connection(engine=None)
        assert conn.engine_name == "graph"
        conn.close()

    def test_explicit_graph_engine(self):
        conn = _make_connection(engine="graph")
        assert conn.engine_name == "graph"
        conn.close()

    def test_engine_mismatch_raises(self):
        transport = httpx.MockTransport(_graph_handler)
        with pytest.raises(BackendOperationError, match="Engine mismatch"):
            ExcelConnection(
                DSN,
                engine="openpyxl",
                credential="tok",
                transport=transport,
            )


class TestGraphConnectionSelect:
    def test_select_all(self):
        conn = _make_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Employees")
        rows = cursor.fetchall()
        assert len(rows) == 3
        assert rows[0] == (1, "Alice", "Eng")
        conn.close()

    def test_select_with_where(self):
        conn = _make_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT name, dept FROM Employees WHERE dept = ?", ("Eng",))
        rows = cursor.fetchall()
        assert len(rows) == 2
        assert rows[0] == ("Alice", "Eng")
        assert rows[1] == ("Carol", "Eng")
        conn.close()

    def test_select_with_order_by(self):
        conn = _make_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM Employees ORDER BY name ASC")
        rows = cursor.fetchall()
        assert rows == [("Alice",), ("Bob",), ("Carol",)]
        conn.close()

    def test_select_with_limit(self):
        conn = _make_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Employees LIMIT 2")
        rows = cursor.fetchall()
        assert len(rows) == 2
        conn.close()


class TestGraphConnectionReadOnlyEnforcement:
    def test_insert_raises(self):
        conn = _make_connection()
        cursor = conn.cursor()
        with pytest.raises(NotSupportedError, match="INSERT"):
            cursor.execute(
                "INSERT INTO Employees (id, name, dept) VALUES (?, ?, ?)",
                (4, "Dave", "HR"),
            )
        conn.close()

    def test_update_raises(self):
        conn = _make_connection()
        cursor = conn.cursor()
        with pytest.raises(NotSupportedError, match="UPDATE"):
            cursor.execute(
                "UPDATE Employees SET dept = ? WHERE name = ?",
                ("HR", "Alice"),
            )
        conn.close()

    def test_delete_raises(self):
        conn = _make_connection()
        cursor = conn.cursor()
        with pytest.raises(NotSupportedError, match="DELETE"):
            cursor.execute("DELETE FROM Employees WHERE id = ?", (1,))
        conn.close()

    def test_create_table_raises(self):
        conn = _make_connection()
        cursor = conn.cursor()
        with pytest.raises(NotSupportedError, match="CREATE"):
            cursor.execute("CREATE TABLE NewSheet (a, b, c)")
        conn.close()

    def test_drop_table_raises(self):
        conn = _make_connection()
        cursor = conn.cursor()
        with pytest.raises(NotSupportedError, match="DROP"):
            cursor.execute("DROP TABLE Employees")
        conn.close()


class TestGraphConnectionModuleConnect:
    def test_connect_function(self):
        transport = httpx.MockTransport(_graph_handler)
        conn = connect(
            DSN,
            engine=None,
            credential="tok",
            transport=transport,
        )
        assert conn.engine_name == "graph"
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Employees")
        assert len(cursor.fetchall()) == 3
        conn.close()


# ---------------------------------------------------------------------------
# v1.3 Transaction guard tests
# ---------------------------------------------------------------------------


class TestGraphConnectionTransactionGuards:
    def test_autocommit_false_rejected(self):
        """Graph backend should reject autocommit=False."""
        handler, _ = _build_handler()
        transport = httpx.MockTransport(handler)
        with pytest.raises(NotSupportedError, match="transactions"):
            ExcelConnection(
                DSN,
                engine=None,
                credential="tok",
                transport=transport,
                autocommit=False,
            )

    def test_autocommit_false_with_readonly_false_rejected(self):
        handler, _ = _build_handler()
        transport = httpx.MockTransport(handler)
        with pytest.raises(NotSupportedError, match="transactions"):
            ExcelConnection(
                DSN,
                engine=None,
                credential="tok",
                transport=transport,
                autocommit=False,
                readonly=False,
            )

    def test_rollback_raises_for_graph(self):
        """rollback() should always raise for non-transactional backend."""
        conn, _ = _make_rw_connection()
        with pytest.raises(NotSupportedError, match="rollback"):
            conn.rollback()
        conn.close()

    def test_commit_is_noop_for_graph(self):
        """commit() should succeed (no-op) for Graph backend."""
        conn, _ = _make_rw_connection()
        conn.commit()  # should not raise
        conn.close()


# ---------------------------------------------------------------------------
# v1.3 Write operation integration tests
# ---------------------------------------------------------------------------


class TestGraphConnectionInsert:
    def test_insert_via_cursor(self):
        conn, state = _make_rw_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO Employees (id, name, dept) VALUES (?, ?, ?)",
            (4, "Dave", "HR"),
        )
        assert cursor.rowcount == 1
        assert cursor.lastrowid is not None
        # Verify PATCH was sent
        patch_reqs = [r for r in state["requests"] if r[0] == "PATCH"]
        assert len(patch_reqs) >= 1
        conn.close()


class TestGraphConnectionUpdate:
    def test_update_via_cursor(self):
        conn, state = _make_rw_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE Employees SET dept = 'HR' WHERE name = 'Alice'",
        )
        assert cursor.rowcount == 1
        # Verify PATCH was sent for write_sheet
        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        assert len(patch_reqs) >= 1
        conn.close()


class TestGraphConnectionDelete:
    def test_delete_via_cursor(self):
        conn, state = _make_rw_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Employees WHERE id = ?", (1,))
        assert cursor.rowcount == 1
        delete_reqs = [
            r for r in state["requests"] if r[0] == "POST" and r[1].endswith("/delete")
        ]
        patch_reqs = [
            r for r in state["requests"] if r[0] == "PATCH" and "/range(" in r[1]
        ]
        assert delete_reqs or patch_reqs
        conn.close()

    def test_delete_all_rows(self):
        conn, state = _make_rw_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Employees")
        assert cursor.rowcount == 3
        conn.close()


class TestGraphConnectionCreateDrop:
    def test_create_table(self):
        conn, state = _make_rw_connection()
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE Products (sku, name, price)")
        # Verify POST to /worksheets/add
        add_reqs = [
            r for r in state["requests"] if r[0] == "POST" and "/worksheets/add" in r[1]
        ]
        created_names = [str(req[2].get("name", "")) for req in add_reqs]
        assert "Products" in created_names
        assert len(add_reqs) >= 1
        conn.close()

    def test_drop_table(self):
        conn, state = _make_rw_connection()
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE TempDrop (id)")
        cursor.execute("DROP TABLE TempDrop")
        delete_reqs = [
            r for r in state["requests"] if r[0] == "DELETE" and "/worksheets/" in r[1]
        ]
        assert len(delete_reqs) == 1
        conn.close()


class TestGraphConnectionResourceLeak:
    """Backend resources are cleaned up if autocommit guard fails."""

    def test_autocommit_false_closes_backend(self):
        """When autocommit=False is rejected, backend.close() should be called."""
        close_called = {"n": 0}
        original_handler, _ = _build_handler()

        def tracking_handler(request: httpx.Request) -> httpx.Response:
            return original_handler(request)

        transport = httpx.MockTransport(tracking_handler)

        # Monkey-patch GraphBackend.close to track calls
        from excel_dbapi.engines.graph.backend import GraphBackend
        original_close = GraphBackend.close

        def counting_close(self: Any) -> None:
            close_called["n"] += 1
            original_close(self)

        GraphBackend.close = counting_close  # type: ignore[assignment]
        try:
            with pytest.raises(NotSupportedError, match="transactions"):
                ExcelConnection(
                    DSN,
                    engine=None,
                    credential="tok",
                    transport=transport,
                    autocommit=False,
                )
            assert close_called["n"] >= 1, "backend.close() was not called on autocommit=False rejection"
        finally:
            GraphBackend.close = original_close  # type: ignore[assignment]

class TestGraphExecutemanySemantics:
    def test_executemany_raises_non_atomic_error_without_restore(self):
        conn, _ = _make_rw_connection()
        cursor = conn.cursor()

        def fail_if_restore_called(snapshot: Any) -> None:
            raise AssertionError("restore() must not be called for graph executemany")

        conn.engine.restore = fail_if_restore_called  # type: ignore[assignment]
        with pytest.raises(
            DatabaseError,
            match="partial writes may have occurred",
        ):
            cursor.executemany(
                "INSERT INTO Employees (id, name, dept) VALUES (?, ?, ?)",
                [(4, "Dan", "HR"), (5, "Eve")],
            )
        conn.close()



def _create_r16_workbook(
    path: Path, sheet: str, headers: list[object], rows: list[list[object]]
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = sheet
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)
    workbook.close()

def test_graph_invalid_credential_is_wrapped_as_operational_error() -> None:
    with pytest.raises(OperationalError, match="Cannot normalise"):
        connect(
            "msgraph://drives/d/items/i",
            engine="graph",
            credential=cast(Any, 42),
        )

def test_graph_token_provider_failure_is_translated_during_execute() -> None:
    class ExplodingTokenProvider:
        def get_token(self, *args: Any) -> Any:
            del args
            raise RuntimeError("token boom")

    transport = httpx.MockTransport(lambda request: httpx.Response(200, json={}))
    with ExcelConnection(
        "msgraph://drives/drv-test/items/itm-test",
        engine="graph",
        credential=ExplodingTokenProvider(),
        transport=transport,
    ) as conn:
        cursor = conn.cursor()
        with pytest.raises(
            Error, match="Failed to acquire authentication token: token boom"
        ):
            cursor.execute("SELECT * FROM Employees")



def test_graph_client_transport_error_and_session_property() -> None:
    class BrokenTransport(httpx.BaseTransport):
        def handle_request(self, request: httpx.Request) -> httpx.Response:
            raise httpx.ConnectError("boom", request=request)

    client = GraphClient(normalize_token_provider("tok"), transport=BrokenTransport())
    assert client.session_id is None
    client.session_id = "sid"
    assert client.session_id == "sid"
    with pytest.raises(OperationalError, match="Graph API request failed"):
        client.get("/x")
    client.close()

def test_graph_session_reopen_and_close_swallow_errors() -> None:
    class FakeClient:
        def __init__(self) -> None:
            self.session_id: str | None = "sid-1"
            self.calls = 0

        def post(self, path: str, json: dict[str, Any]) -> Any:
            self.calls += 1
            if path.endswith("closeSession"):
                raise RuntimeError("close fail")
            return types.SimpleNamespace(json=lambda: {"id": "sid-2"})

    class FakeLocator:
        item_path = "/drives/d/items/i"

    client = FakeClient()
    session = WorkbookSession(
        cast(Any, client), cast(Any, FakeLocator()), persist_changes=False
    )
    session._open = True
    session.reopen()
    assert session.is_open is True
    session.close()
    assert session.is_open is False

def test_graph_auth_additional_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    class BadSignatureObj:
        def get_token(self) -> str:
            return "x"

    def explode_signature(obj: Any) -> Any:
        raise TypeError("no signature")

    monkeypatch.setattr("inspect.signature", explode_signature)
    assert _has_get_token_with_args(BadSignatureObj()) is False

    module_azure = types.ModuleType("azure")
    module_identity = types.ModuleType("azure.identity")

    class DefaultAzureCredential:
        def get_token(self, scope: str) -> Any:
            return types.SimpleNamespace(token=f"azure:{scope}")

    setattr(module_identity, "DefaultAzureCredential", DefaultAzureCredential)
    setattr(module_azure, "identity", module_identity)
    monkeypatch.setitem(sys.modules, "azure", module_azure)
    monkeypatch.setitem(sys.modules, "azure.identity", module_identity)
    provider = normalize_token_provider(None)
    assert provider.get_token().startswith("azure:")
