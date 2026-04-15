from typing import Any

import pytest

from excel_dbapi.engines.base import TableData, WorkbookBackend
from excel_dbapi.exceptions import NotSupportedError


class DummyBackend(WorkbookBackend):
    @property
    def readonly(self) -> bool:
        return False

    @property
    def supports_transactions(self) -> bool:
        return True

    def load(self) -> None:
        pass

    def save(self) -> None:
        pass

    def snapshot(self) -> Any:
        return {"state": "ok"}

    def restore(self, snapshot: Any) -> None:
        pass

    def list_sheets(self) -> list[str]:
        return ["Sheet1"]

    def read_sheet(self, sheet_name: str) -> TableData:
        return TableData(headers=["id"], rows=[[1]])

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        pass

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        return 2

    def create_sheet(self, name: str, headers: list[str]) -> None:
        pass

    def drop_sheet(self, name: str) -> None:
        pass


def test_default_get_workbook_raises_not_supported() -> None:
    backend = DummyBackend("dummy.xlsx")
    with pytest.raises(NotSupportedError):
        backend.get_workbook()


def test_workbook_backend_methods_coverage() -> None:
    backend = DummyBackend("dummy.xlsx")
    backend.load()
    backend.save()
    backend.snapshot()
    backend.restore({})
    assert backend.list_sheets() == ["Sheet1"]
    assert backend.read_sheet("Sheet1").headers == ["id"]
    assert backend.append_row("Sheet1", [2]) == 2
