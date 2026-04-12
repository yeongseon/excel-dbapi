from collections.abc import Callable
from functools import wraps
from typing import Any, Concatenate, List, Optional, ParamSpec, TypeVar, cast

from .engines.result import Description, ExecutionResult
from .exceptions import InterfaceError, NotSupportedError, ProgrammingError


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

    def __init__(self, connection: Any):
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
        self._results = result.rows
        self._index = 0
        self.description = result.description
        if not self.description:
            self.description = None
        self.rowcount = result.rowcount
        self.lastrowid = result.lastrowid
        if self.connection.autocommit and result.action in {
            "INSERT",
            "CREATE",
            "DROP",
            "UPDATE",
            "DELETE",
        }:
            self.connection.engine.save()
            self.connection._snapshot = self.connection.engine.snapshot()
        return self

    @check_closed
    def executemany(
        self, query: str, seq_of_params: List[tuple[Any, ...]]
    ) -> "ExcelCursor":
        total_rowcount = 0
        last_rowid = None
        last_action = None
        snapshot = None
        if not self.connection.autocommit:
            snapshot = self.connection.engine.snapshot()
        for params in seq_of_params:
            try:
                result: ExecutionResult = self.connection.execute(query, params)
            except ValueError as exc:
                if snapshot is not None:
                    self.connection.engine.restore(snapshot)
                raise ProgrammingError(str(exc)) from exc
            except NotImplementedError as exc:
                if snapshot is not None:
                    self.connection.engine.restore(snapshot)
                raise NotSupportedError(str(exc)) from exc
            total_rowcount += result.rowcount
            last_rowid = result.lastrowid
            last_action = result.action
        self._results = []
        self._index = 0
        self.description = None
        self.rowcount = total_rowcount
        self.lastrowid = last_rowid
        if self.connection.autocommit and last_action in {
            "INSERT",
            "CREATE",
            "DROP",
            "UPDATE",
            "DELETE",
        }:
            self.connection.engine.save()
            self.connection._snapshot = self.connection.engine.snapshot()
        return self

    @check_closed
    def fetchone(self) -> Optional[tuple[Any, ...]]:
        if self._index >= len(self._results):
            return None
        result = self._results[self._index]
        self._index += 1
        return result

    @check_closed
    def fetchall(self) -> List[tuple[Any, ...]]:
        results = self._results[self._index :]
        self._index = len(self._results)
        return results

    @check_closed
    def fetchmany(self, size: Optional[int] = None) -> List[tuple[Any, ...]]:
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
