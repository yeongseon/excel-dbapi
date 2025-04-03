from typing import Any, Dict, List, Optional

from .engine.executor import execute_query
from .engine.parser import parse_sql
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

        Args:
            connection (ExcelConnection): The parent connection object.
        """
        self.connection = connection
        self.closed: bool = False
        self._results: List[Dict[str, Any]] = []
        self._index: int = 0

    @check_closed
    def execute(self, query: str, params: Optional[tuple] = None) -> "ExcelCursor":
        """
        Execute a SQL query.

        Args:
            query (str): The SQL query string.
            params (Optional[tuple]): Parameters to bind to query placeholders.

        Returns:
            ExcelCursor: The cursor itself.
        """
        parsed = parse_sql(query, params)
        self._results = execute_query(parsed, self.connection.data)
        self._index = 0
        return self

    @check_closed
    def fetchone(self) -> Optional[Dict[str, Any]]:
        """
        Fetch the next row of a query result.

        Returns:
            Optional[Dict[str, Any]]: The next row or None if no more rows.
        """
        if self._index >= len(self._results):
            return None
        result = self._results[self._index]
        self._index += 1
        return result

    @check_closed
    def fetchall(self) -> List[Dict[str, Any]]:
        """
        Fetch all remaining rows of a query result.

        Returns:
            List[Dict[str, Any]]: List of all remaining rows.
        """
        results = self._results[self._index :]
        self._index = len(self._results)
        return results

    def close(self) -> None:
        """
        Close the cursor.
        """
        self.closed = True
