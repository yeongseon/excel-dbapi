from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from functools import wraps
from typing import (
    TYPE_CHECKING,
    Any,
    Concatenate,
    List,
    Optional,
    ParamSpec,
    TypeVar,
    cast,
)

from .engines.result import Description, ExecutionResult

if TYPE_CHECKING:
    from .connection import ExcelConnection
from .exceptions import (
    DatabaseError,
    InterfaceError,
    NotSupportedError,
    OperationalError,
    ProgrammingError,
)


P = ParamSpec("P")
R = TypeVar("R")


def check_closed(
    func: Callable[Concatenate["ExcelCursor", P], R],
) -> Callable[Concatenate["ExcelCursor", P], R]:
    """Decorator to check if cursor is closed before executing method."""

    @wraps(func)
    def wrapper(self: "ExcelCursor", *args: P.args, **kwargs: P.kwargs) -> R:
        if self.closed:
            raise InterfaceError("Cursor is already closed")
        if self.connection.closed:
            raise InterfaceError("Cannot operate on a closed connection")
        return func(self, *args, **kwargs)

    return cast(Callable[Concatenate["ExcelCursor", P], R], wrapper)

def _map_exception(exc: Exception) -> DatabaseError:
    """Map a non-DB-API exception to the appropriate DB-API exception type."""
    if isinstance(exc, ValueError):
        return ProgrammingError(str(exc))
    if isinstance(exc, NotImplementedError):
        return NotSupportedError(str(exc))
    if isinstance(exc, (KeyError, TypeError, IndexError)):
        return ProgrammingError(str(exc))
    if isinstance(exc, OSError):
        return OperationalError(str(exc))
    return DatabaseError(str(exc))


class ExcelCursor:
    """
    ExcelCursor provides a PEP 249 compliant Cursor interface
    for executing SQL-like queries on Excel data.
    """

    def __init__(self, connection: ExcelConnection):
        """
        Initialize the cursor with a connection.
        """
        self.connection = connection
        self.closed: bool = False
        self._results: List[tuple[Any, ...]] = []
        self._index: int = 0
        self.description: Description | None = None
        self.rowcount = -1
        self.lastrowid: int | None = None
        self.arraysize = 1
        self._has_result_set = False

    def _reset_state(self) -> None:
        """Reset cursor state to defaults after an error."""
        self._results = []
        self._index = 0
        self.description = None
        self.rowcount = -1
        self.lastrowid = None
        self._has_result_set = False


    @check_closed
    def execute(self, query: str, params: Sequence[Any] | None = None) -> "ExcelCursor":
        self._reset_state()
        try:
            result: ExecutionResult = self.connection.execute(query, params)
        except DatabaseError:
            raise
        except Exception as exc:
            raise _map_exception(exc) from exc
        self._results = result.rows
        self._index = 0
        self.description = result.description
        if not self.description:
            self.description = None
        self._has_result_set = result.action in {"SELECT", "COMPOUND"}
        self.rowcount = result.rowcount
        self.lastrowid = result.lastrowid
        return self

    @check_closed
    def executemany(
        self, query: str, seq_of_params: Iterable[Sequence[Any]]
    ) -> "ExcelCursor":
        self._reset_state()
        ensure_write_lock = getattr(
            self.connection, "_ensure_write_lock_for_query", None
        )
        if callable(ensure_write_lock):
            ensure_write_lock(query)
        total_rowcount = 0
        last_rowid = None
        last_action = None
        supports_transactions = bool(
            getattr(self.connection.engine, "supports_transactions", True)
        )
        snapshot = self.connection.engine.snapshot() if supports_transactions else None
        backend_name = type(self.connection.engine).__name__

        for params in seq_of_params:
            try:
                result: ExecutionResult = self.connection._executor.execute_with_params(
                    query, tuple(params)
                )
            except DatabaseError as exc:
                self._reset_state()
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise
                raise type(exc)(
                    f"{exc}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except Exception as exc:
                self._reset_state()
                mapped = _map_exception(exc)
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise mapped from exc
                raise type(mapped)(
                    f"{mapped}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            total_rowcount += result.rowcount
            last_rowid = result.lastrowid
            last_action = result.action
        self._results = []
        self._index = 0
        self.description = None
        self._has_result_set = False
        self.rowcount = total_rowcount
        self.lastrowid = last_rowid
        if self.connection.autocommit and last_action is not None:
            try:
                self.connection._finalize_autocommit(last_action)
            except Exception as exc:
                from excel_dbapi.exceptions import Error as _DBAPIError

                if isinstance(exc, _DBAPIError):
                    raise
                raise OperationalError(str(exc)) from exc
        return self

    @check_closed
    def fetchone(self) -> Optional[tuple[Any, ...]]:
        if not self._has_result_set:
            raise ProgrammingError(
                "No result set: call execute() with a SELECT statement first"
            )
        if self._index >= len(self._results):
            return None
        result = self._results[self._index]
        self._index += 1
        return result

    @check_closed
    def fetchall(self) -> List[tuple[Any, ...]]:
        if not self._has_result_set:
            raise ProgrammingError(
                "No result set: call execute() with a SELECT statement first"
            )
        results = self._results[self._index :]
        self._index = len(self._results)
        return results

    @check_closed
    def fetchmany(self, size: Optional[int] = None) -> List[tuple[Any, ...]]:
        if not self._has_result_set:
            raise ProgrammingError(
                "No result set: call execute() with a SELECT statement first"
            )
        count = self.arraysize if size is None else size
        if count <= 0:
            return []
        start = self._index
        end = min(self._index + count, len(self._results))
        self._index = end
        return self._results[start:end]

    def close(self) -> None:
        self.closed = True

    @check_closed
    def setinputsizes(self, sizes: Any) -> None:
        """PEP 249 no-op.  Accepted for compatibility."""
        pass

    @check_closed
    def setoutputsize(self, size: int, column: Optional[int] = None) -> None:
        """PEP 249 no-op.  Accepted for compatibility."""
        pass
