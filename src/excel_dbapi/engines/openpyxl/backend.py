from io import BytesIO
import os
import sys
import tempfile
from typing import Any, cast

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from ...exceptions import BackendOperationError
from ...executor import SharedExecutor
from ..result import ExecutionResult
from ..base import TableData, WorkbookBackend, _normalize_headers


class OpenpyxlBackend(WorkbookBackend):

    @property
    def readonly(self) -> bool:
        return False

    @property
    def supports_transactions(self) -> bool:
        return True

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
            raise BackendOperationError("Workbook is not loaded")
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
            raise BackendOperationError("Workbook is not loaded")
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
            raise BackendOperationError(f"Sheet '{sheet_name}' not found in Excel")
        row_iter = ws.iter_rows(values_only=True)
        first_row = next(row_iter, None)
        if first_row is None:
            return TableData(headers=[], rows=[])

        # Trim trailing None/empty columns left by in-place column deletion.
        raw_headers = list(first_row)
        while raw_headers and (raw_headers[-1] is None or (isinstance(raw_headers[-1], str) and raw_headers[-1].strip() == "")):
            raw_headers.pop()
        if not raw_headers:
            return TableData(headers=[], rows=[])
        num_cols = len(raw_headers)
        headers = _normalize_headers(raw_headers)
        table_rows: list[list[Any]] = []
        approx_bytes = sys.getsizeof(headers)

        for index, row in enumerate(row_iter, start=1):
            row_values = list(row)[:num_cols]
            table_rows.append(row_values)
            self._check_row_limit(sheet_name, index)
            approx_bytes += sys.getsizeof(row_values)
            approx_bytes += sum(sys.getsizeof(value) for value in row_values)
            self._check_memory_limit(sheet_name, approx_bytes)

        return TableData(headers=headers, rows=table_rows)

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        ws = self.data.get(sheet_name)
        if ws is None:
            raise BackendOperationError(f"Sheet '{sheet_name}' not found in Excel")
        # Write in-place to preserve cell formatting (fonts, borders, fills).
        # Step 1: Write header row.
        for col_idx, header in enumerate(data.headers, start=1):
            ws.cell(row=1, column=col_idx).value = header
        # Clear extra header columns if the new header is narrower.
        old_max_col = ws.max_column
        for col_idx in range(len(data.headers) + 1, old_max_col + 1):
            ws.cell(row=1, column=col_idx).value = None
        # Step 2: Write data rows in-place.
        for row_idx, row in enumerate(data.rows, start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx).value = value
            # Clear extra columns in this row.
            for col_idx in range(len(row) + 1, old_max_col + 1):
                ws.cell(row=row_idx, column=col_idx).value = None
        # Step 3: Remove surplus rows (if data shrunk).
        new_max_row = len(data.rows) + 1  # +1 for header
        if ws.max_row > new_max_row:
            ws.delete_rows(new_max_row + 1, ws.max_row - new_max_row)

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        ws = self.data.get(sheet_name)
        if ws is None:
            raise BackendOperationError(f"Sheet '{sheet_name}' not found in Excel")
        ws.append(row)
        return cast(int, ws.max_row)

    def create_sheet(self, name: str, headers: list[str]) -> None:
        if self.workbook is None:
            raise BackendOperationError("Workbook is not loaded")
        if name in self.data:
            raise BackendOperationError(f"Sheet '{name}' already exists")
        ws = self.workbook.create_sheet(title=name)
        ws.append(headers)
        self.data[name] = ws

    def drop_sheet(self, name: str) -> None:
        if self.workbook is None:
            raise BackendOperationError("Workbook is not loaded")
        ws = self.data.get(name)
        if ws is None:
            raise BackendOperationError(f"Sheet '{name}' not found in Excel")
        self.workbook.remove(ws)
        del self.data[name]

    def get_workbook(self) -> Any:
        if self.workbook is None:
            raise BackendOperationError("Workbook is not loaded")
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
