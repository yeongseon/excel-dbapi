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
    Error,
    InterfaceError,
    ProgrammingError,
    map_exception,
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
        except Error:
            raise
        except Exception as exc:
            raise map_exception(exc) from exc
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
        try:
            result: ExecutionResult = self.connection.executemany(query, seq_of_params)
        except Error:
            raise
        except Exception as exc:
            raise map_exception(exc) from exc
        self._results = []
        self._index = 0
        self.description = None
        self._has_result_set = False
        self.rowcount = result.rowcount
        self.lastrowid = result.lastrowid
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
