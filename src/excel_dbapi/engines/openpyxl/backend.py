from io import BytesIO
import os
import sys
import tempfile
from typing import Any, cast

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from ...executor import SharedExecutor
from ..result import ExecutionResult
from ..base import TableData, WorkbookBackend, _normalize_headers


class OpenpyxlBackend(WorkbookBackend):
    def __init__(
        self,
        file_path: str,
        *,
        data_only: bool = True,
        create: bool = False,
        sanitize_formulas: bool = True,
        **options: Any,
    ) -> None:
        super().__init__(
            file_path,
            data_only=data_only,
            create=create,
            sanitize_formulas=sanitize_formulas,
            **options,
        )
        self._data_only = data_only
        self.workbook: Workbook | None = None
        self.data: dict[str, Any] = {}
        self.load()

    def load(self) -> None:
        if self.create and (
            not os.path.exists(self.file_path) or os.path.getsize(self.file_path) == 0
        ):
            self.workbook = Workbook()
            self.workbook.save(self.file_path)
        else:
            self.workbook = load_workbook(self.file_path, data_only=self._data_only)
        self.data = {sheet: self.workbook[sheet] for sheet in self.workbook.sheetnames}

    def save(self) -> None:
        if self.workbook is None:
            raise ValueError("Workbook is not loaded")
        directory = os.path.dirname(self.file_path) or "."
        temp_file = None
        try:
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".xlsx", dir=directory
            ) as handle:
                temp_file = handle.name
            os.chmod(temp_file, 0o600)
            self.workbook.save(temp_file)
            os.replace(temp_file, self.file_path)
        finally:
            if temp_file and os.path.exists(temp_file):
                os.unlink(temp_file)

    def snapshot(self) -> BytesIO:
        if self.workbook is None:
            raise ValueError("Workbook is not loaded")
        buffer = BytesIO()
        self.workbook.save(buffer)
        buffer.seek(0)
        return buffer

    def restore(self, snapshot: Any) -> None:
        snapshot.seek(0)
        self.workbook = load_workbook(snapshot, data_only=self._data_only)
        self.data = {sheet: self.workbook[sheet] for sheet in self.workbook.sheetnames}

    def list_sheets(self) -> list[str]:
        return list(self.data.keys())

    def read_sheet(self, sheet_name: str) -> TableData:
        ws = self.data.get(sheet_name)
        if ws is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")
        row_iter = ws.iter_rows(values_only=True)
        first_row = next(row_iter, None)
        if first_row is None:
            return TableData(headers=[], rows=[])

        headers = _normalize_headers(list(first_row))
        table_rows: list[list[Any]] = []
        approx_bytes = sys.getsizeof(headers)

        for index, row in enumerate(row_iter, start=1):
            row_values = list(row)
            table_rows.append(row_values)
            self._check_row_limit(sheet_name, index)
            approx_bytes += sys.getsizeof(row_values)
            approx_bytes += sum(sys.getsizeof(value) for value in row_values)
            self._check_memory_limit(sheet_name, approx_bytes)

        return TableData(headers=headers, rows=table_rows)

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        ws = self.data.get(sheet_name)
        if ws is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")
        ws.delete_rows(1, ws.max_row)
        ws.append(data.headers)
        for row in data.rows:
            ws.append(row)

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        ws = self.data.get(sheet_name)
        if ws is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")
        ws.append(row)
        return cast(int, ws.max_row)

    def create_sheet(self, name: str, headers: list[str]) -> None:
        if self.workbook is None:
            raise ValueError("Workbook is not loaded")
        if name in self.data:
            raise ValueError(f"Sheet '{name}' already exists")
        ws = self.workbook.create_sheet(title=name)
        ws.append(headers)
        self.data[name] = ws

    def drop_sheet(self, name: str) -> None:
        if self.workbook is None:
            raise ValueError("Workbook is not loaded")
        ws = self.data.get(name)
        if ws is None:
            raise ValueError(f"Sheet '{name}' not found in Excel")
        self.workbook.remove(ws)
        del self.data[name]

    def get_workbook(self) -> Any:
        if self.workbook is None:
            raise ValueError("Workbook is not loaded")
        return self.workbook

    def execute(self, query: str) -> ExecutionResult:
        return SharedExecutor(
            self, sanitize_formulas=self.sanitize_formulas
        ).execute_with_params(query, None)

    def execute_with_params(
        self, query: str, params: tuple[Any, ...] | None = None
    ) -> ExecutionResult:
        return SharedExecutor(
            self, sanitize_formulas=self.sanitize_formulas
        ).execute_with_params(query, params)
