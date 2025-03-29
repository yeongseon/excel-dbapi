from typing import Callable, Dict, List, Optional, Union

from .exceptions import OperationalError
from .table import ExcelTable


class QueryEngine:
    """Engine for executing SQL-like queries on Excel tables."""

    def __init__(self, table: ExcelTable) -> None:
        """Initialize with a table instance."""
        self.table = table

    def select(
        self,
        columns: Optional[List[str]] = None,
        where: Optional[Callable[[List[Union[str, int, float]]], bool]] = None,
        limit: Optional[int] = None,
    ) -> List[Dict[str, Union[str, int, float]]]:
        """Execute a SELECT query with optional columns, where clause, and limit."""
        if not self.table.sheet:
            raise OperationalError("Table must be opened before querying.")

        headers = [cell.value for cell in self.table.sheet[1]]
        if not headers or not all(isinstance(h, str) for h in headers):
            raise OperationalError("First row must contain valid string headers.")

        if columns:
            col_indices = [headers.index(col) for col in columns if col in headers]
        else:
            col_indices = list(range(len(headers)))

        rows = self.table.fetch_all()[1:]
        result: List[Dict[str, Union[str, int, float]]] = []

        for row in rows:
            if where is None or where(row):
                row_dict = {headers[i]: row[i] for i in col_indices}
                result.append(row_dict)
                if limit and len(result) >= limit:
                    break

        return result


def query(
    table: ExcelTable,
    columns: Optional[List[str]] = None,
    where: Optional[Callable[[List[Union[str, int, float]]], bool]] = None,
    limit: Optional[int] = None,
) -> List[Dict[str, Union[str, int, float]]]:
    """Convenience function to execute a query."""
    engine = QueryEngine(table)
    return engine.select(columns=columns, where=where, limit=limit)
