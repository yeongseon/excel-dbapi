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
        return func(self, *args, **kwargs)

    return cast(Callable[Concatenate["ExcelCursor", P], R], wrapper)


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

    @check_closed
    def execute(
        self, query: str, params: Optional[tuple[Any, ...]] = None
    ) -> "ExcelCursor":
        try:
            result: ExecutionResult = self.connection.execute(query, params)
        except ValueError as exc:
            raise ProgrammingError(str(exc)) from exc
        except NotImplementedError as exc:
            raise NotSupportedError(str(exc)) from exc
        except (KeyError, TypeError, IndexError) as exc:
            raise ProgrammingError(str(exc)) from exc
        except OSError as exc:
            raise OperationalError(str(exc)) from exc
        except DatabaseError:
            raise
        except Exception as exc:
            raise DatabaseError(str(exc)) from exc
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
            except ValueError as exc:
                mapped: Exception = ProgrammingError(str(exc))
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise mapped from exc
                raise ProgrammingError(
                    f"{mapped}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except NotImplementedError as exc:
                mapped = NotSupportedError(str(exc))
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise mapped from exc
                raise NotSupportedError(
                    f"{mapped}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except (KeyError, TypeError, IndexError) as exc:
                mapped = ProgrammingError(str(exc))
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise mapped from exc
                raise ProgrammingError(
                    f"{mapped}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except OSError as exc:
                mapped = OperationalError(str(exc))
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise mapped from exc
                raise OperationalError(
                    f"{mapped}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except DatabaseError as exc:
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise
                raise type(exc)(
                    f"{exc}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except Exception as exc:
                mapped = DatabaseError(str(exc))
                if supports_transactions:
                    self.connection.engine.restore(snapshot)
                    raise mapped from exc
                raise DatabaseError(
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
            self.connection._finalize_autocommit(last_action)
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
