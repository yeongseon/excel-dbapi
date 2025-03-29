from typing import Dict, List, Optional, Union

from .connection import ExcelConnection
from .exceptions import OperationalError, ProgrammingError
from .query import QueryEngine
from .table import ExcelTable


class ExcelCursor:
    """DBAPI-compliant cursor for executing queries on Excel tables."""

    def __init__(self, connection: ExcelConnection) -> None:
        """Initialize with a connection."""
        self.connection = connection
        self.table: Optional[ExcelTable] = None
        self._results: Optional[List[Dict[str, Union[str, int, float]]]] = None
        self._rowcount = -1

    def execute(self, query: str, params: Optional[tuple] = None) -> None:
        """Execute a SQL-like query."""
        if not query.strip().upper().startswith("SELECT"):
            raise ProgrammingError("Only SELECT queries are supported.")

        if "FROM" not in query.upper():
            raise ProgrammingError("Query must include FROM clause.")

        table_name = query.split("FROM", 1)[1].strip().split()[0]
        self.table = ExcelTable(self.connection, table_name)

        with self.table:
            qe = QueryEngine(self.table)
            if "*" in query:
                self._results = qe.select()
            else:
                columns = [
                    col.strip()
                    for col in query.split("SELECT")[1].split("FROM")[0].split(",")
                ]
                self._results = qe.select(columns=columns)
            self._rowcount = len(self._results or [])

    def fetchone(self) -> Optional[Dict[str, Union[str, int, float]]]:
        """Fetch the next row of the query result."""
        if self._results is None:
            raise OperationalError("No query executed.")
        if not self._results:
            return None
        return self._results.pop(0)

    def fetchall(self) -> List[Dict[str, Union[str, int, float]]]:
        """Fetch all rows of the query result."""
        if self._results is None:
            raise OperationalError("No query executed.")
        results = self._results or []
        self._results = []
        return results

    def fetchmany(self, size: int = 1) -> List[Dict[str, Union[str, int, float]]]:
        """Fetch a specified number of rows."""
        if self._results is None:
            raise OperationalError("No query executed.")
        results = self._results[:size]
        self._results = self._results[size:]
        return results

    @property
    def rowcount(self) -> int:
        """Return the number of rows affected."""
        return self._rowcount

    def close(self) -> None:
        """Close the cursor."""
        self.table = None
        self._results = None
        self._rowcount = -1
