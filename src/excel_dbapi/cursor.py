from typing import Any, List, Optional

from .engine.result import ExecutionResult
from .exceptions import InterfaceError


def check_closed(func):
    """Decorator to check if cursor is closed before executing method."""
    def wrapper(self, *args, **kwargs):
        if self.closed:
            raise InterfaceError("Cursor is already closed")
        return func(self, *args, **kwargs)
    return wrapper


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
        self._results: List[tuple] = []
        self._index: int = 0
        self.description = None
        self.rowcount = -1
        self.lastrowid = None
        self.arraysize = 1

    @check_closed
    def execute(self, query: str, params: Optional[tuple] = None) -> "ExcelCursor":
        result: ExecutionResult = self.connection.engine.execute_with_params(query, params)
        self._results = result.rows
        self._index = 0
        self.description = result.description
        self.rowcount = result.rowcount
        self.lastrowid = result.lastrowid
        return self

    @check_closed
    def fetchone(self) -> Optional[tuple]:
        if self._index >= len(self._results):
            return None
        result = self._results[self._index]
        self._index += 1
        return result

    @check_closed
    def fetchall(self) -> List[tuple]:
        results = self._results[self._index:]
        self._index = len(self._results)
        return results

    @check_closed
    def fetchmany(self, size: Optional[int] = None) -> List[tuple]:
        count = self.arraysize if size is None else size
        if count <= 0:
            return []
        start = self._index
        end = min(self._index + count, len(self._results))
        self._index = end
        return self._results[start:end]

    def close(self) -> None:
        self.closed = True
