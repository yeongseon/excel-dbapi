from io import BytesIO
from typing import Optional, Union

import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..exceptions import OperationalError


class OpenPyXLEngine:
    """Excel processing engine using openpyxl."""

    def __init__(self) -> None:
        """Initialize the engine."""
        self._workbook: Optional[Workbook] = None

    def load_workbook(self, file_path: Union[str, BytesIO]) -> "OpenPyXLEngine":
        """Load the workbook from a file path or BytesIO."""
        try:
            self._workbook = openpyxl.load_workbook(file_path)
            return self
        except Exception as e:
            raise OperationalError(f"Failed to load workbook: {e}")

    def close(self) -> None:
        """Close the workbook."""
        if self._workbook:
            self._workbook.close()
            self._workbook = None

    def get_sheet(self, sheet_name: str) -> Worksheet:
        """Get a specific sheet by name."""
        if not self._workbook:
            raise OperationalError("Workbook not loaded.")
        try:
            return self._workbook[sheet_name]
        except KeyError:
            raise OperationalError(f"Sheet '{sheet_name}' not found.")

    @property
    def workbook(self) -> Workbook:
        """Property to access the loaded workbook."""
        if not self._workbook:
            raise OperationalError("Workbook not loaded.")
        return self._workbook
