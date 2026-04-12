from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any


@dataclass
class TableData:
    headers: list[str]
    rows: list[list[Any]]


class WorkbookBackend(ABC):
    file_path: str
    create: bool
    sanitize_formulas: bool
    readonly: bool = False
    supports_transactions: bool = True

    def __init__(
        self,
        file_path: str,
        *,
        data_only: bool = True,
        create: bool = False,
        sanitize_formulas: bool = True,
        **options: Any,
    ) -> None:
        self.file_path = file_path
        self.create = create
        self.sanitize_formulas = sanitize_formulas

    @abstractmethod
    def load(self) -> None:
        pass

    @abstractmethod
    def save(self) -> None:
        pass

    @abstractmethod
    def snapshot(self) -> Any:
        pass

    @abstractmethod
    def restore(self, snapshot: Any) -> None:
        pass

    @abstractmethod
    def list_sheets(self) -> list[str]:
        pass

    @abstractmethod
    def read_sheet(self, sheet_name: str) -> TableData:
        pass

    @abstractmethod
    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        pass

    @abstractmethod
    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        pass

    @abstractmethod
    def create_sheet(self, name: str, headers: list[str]) -> None:
        pass

    @abstractmethod
    def drop_sheet(self, name: str) -> None:
        pass

    def close(self) -> None:
        """Release backend resources.  Default is a no-op."""

    def get_workbook(self) -> Any:
        from ..exceptions import NotSupportedError

        raise NotSupportedError(
            f"Backend '{type(self).__name__}' does not expose a workbook object"
        )
