from typing import Any, Callable, Dict, List, Optional, Union

from .connection import ExcelConnection
from .exceptions import OperationalError


class ExcelTable:
    """Class to treat an Excel sheet as a table."""

    def __init__(self, connection: ExcelConnection, sheet_name: str) -> None:
        """Initialize with a connection and sheet name."""
        self.connection = connection
        self.sheet_name = sheet_name
        self.sheet: Optional[Any] = None

    def open(self) -> "ExcelTable":
        """Open the sheet and prepare it as a table."""
        try:
            self.sheet = self.connection.engine.get_sheet(self.sheet_name)
            return self
        except Exception as e:
            raise OperationalError(f"Failed to open sheet '{self.sheet_name}': {e}")

    def fetch_all(self) -> List[List[Union[str, int, float]]]:
        """Fetch all data from the sheet."""
        if not self.sheet:
            raise OperationalError("Table is not opened.")
        return [[cell.value for cell in row] for row in self.sheet.rows]

    def fetch_row(self, row_num: int) -> List[Union[str, int, float]]:
        """Fetch a specific row by number (1-based index)."""
        if not self.sheet:
            raise OperationalError("Table is not opened.")
        return [cell.value for cell in self.sheet[row_num]]

    def query(
        self,
        columns: Optional[List[str]] = None,
        where: Optional[Callable[[List[Union[str, int, float]]], bool]] = None,
        limit: Optional[int] = None,
    ) -> List[Dict[str, Union[str, int, float]]]:
        """Execute a query on the table."""
        from .query import query  # lazy import to avoid circular dependency

        return query(self, columns=columns, where=where, limit=limit)

    def __enter__(self) -> "ExcelTable":
        """Context manager entry."""
        return self.open()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.sheet = None
