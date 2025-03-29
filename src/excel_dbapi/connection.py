from io import BytesIO
from pathlib import Path
from typing import Optional, Protocol, Union

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .engines.openpyxl_engine import OpenPyXLEngine
from .exceptions import OperationalError
from .remote import fetch_remote_file


class ExcelEngine(Protocol):
    """Protocol for Excel processing engines."""

    def load_workbook(self, file_path: Union[str, BytesIO]) -> "ExcelEngine": ...
    def close(self) -> None: ...
    def get_sheet(self, sheet_name: str) -> Worksheet: ...
    @property
    def workbook(self) -> Workbook: ...


class ExcelConnection:
    """DBAPI-compliant connection to an Excel file."""

    def __init__(
        self, file_path: Union[str, Path], engine: Optional[ExcelEngine] = None
    ) -> None:
        """Initialize with file path and optional engine."""
        self.file_path = Path(file_path)
        self.engine: ExcelEngine = engine if engine else OpenPyXLEngine()
        self._connected = False

    def connect(self) -> "ExcelConnection":
        """Establish connection to the Excel file."""
        try:
            if str(self.file_path).startswith(("http://", "https://")):
                file_data = fetch_remote_file(self.file_path)
                self.engine.load_workbook(file_data)
            else:
                self.engine.load_workbook(str(self.file_path))
            self._connected = True
            return self
        except Exception as e:
            raise OperationalError(f"Failed to connect to {self.file_path}: {e}")

    def close(self) -> None:
        """Close the connection."""
        if self._connected:
            self.engine.close()
            self._connected = False

    def commit(self) -> None:
        """Commit changes (not applicable for read-only Excel, placeholder)."""
        pass

    def rollback(self) -> None:
        """Rollback changes (not applicable, placeholder)."""
        pass

    def cursor(self):  # 타입 주석 제거 및 임포트 지연
        """Return a new cursor object."""
        from .cursor import ExcelCursor  # 메서드 안에서 임포트

        if not self._connected:
            raise OperationalError("Connection is not established.")
        return ExcelCursor(self)

    def __enter__(self) -> "ExcelConnection":
        """Context manager entry."""
        return self.connect()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.close()
