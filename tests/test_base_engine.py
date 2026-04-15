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



class MemoryBackend(WorkbookBackend):
    def __init__(self, sheets: dict[str, TableData], readonly: bool = False) -> None:
        super().__init__("memory.xlsx")
        self._sheets = sheets
        self._readonly = readonly

    @property
    def readonly(self) -> bool:
        return self._readonly

    @property
    def supports_transactions(self) -> bool:
        return True


    def load(self) -> None:
        return None

    def save(self) -> None:
        return None

    def snapshot(self) -> dict[str, TableData]:
        return self._sheets

    def restore(self, snapshot: Any) -> None:
        self._sheets = snapshot

    def list_sheets(self) -> list[str]:
        return list(self._sheets.keys())

    def read_sheet(self, sheet_name: str) -> TableData:
        if sheet_name not in self._sheets:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")
        return self._sheets[sheet_name]

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        self._sheets[sheet_name] = data

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        self._sheets[sheet_name].rows.append(row)
        return len(self._sheets[sheet_name].rows) + 1

    def create_sheet(self, name: str, headers: list[str]) -> None:
        self._sheets[name] = TableData(headers=headers, rows=[])

    def drop_sheet(self, name: str) -> None:
        del self._sheets[name]

def test_base_abstract_method_bodies_are_executable() -> None:
    backend = MemoryBackend({"T": TableData(headers=["id"], rows=[[1]])})
    WorkbookBackend.load(backend)
    WorkbookBackend.save(backend)
    WorkbookBackend.snapshot(backend)
    WorkbookBackend.restore(backend, None)
    WorkbookBackend.list_sheets(backend)
    WorkbookBackend.read_sheet(backend, "T")
    WorkbookBackend.write_sheet(backend, "T", TableData(headers=["id"], rows=[]))
    WorkbookBackend.append_row(backend, "T", [2])
    WorkbookBackend.create_sheet(backend, "U", ["id"])
    WorkbookBackend.drop_sheet(backend, "U")
