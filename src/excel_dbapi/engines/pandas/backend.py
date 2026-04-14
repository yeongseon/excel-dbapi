from typing import Any
import os
import re
import tempfile

import pandas as pd

from ...exceptions import DataError, NotSupportedError
from ...executor import SharedExecutor
from ..base import TableData, WorkbookBackend, _normalize_headers
from ..result import ExecutionResult


class PandasBackend(WorkbookBackend):
    def __init__(
        self,
        file_path: str,
        *,
        data_only: bool = True,
        create: bool = False,
        sanitize_formulas: bool = True,
        **options: Any,
    ) -> None:
        if not data_only:
            raise NotSupportedError(
                "The pandas backend does not support data_only=False; use the openpyxl backend instead"
            )
        super().__init__(
            file_path,
            data_only=data_only,
            create=create,
            sanitize_formulas=sanitize_formulas,
            **options,
        )
        self._data_only = data_only
        self.data: dict[str, pd.DataFrame] = {}
        self.load()

    def load(self) -> None:
        if self.create and (
            not os.path.exists(self.file_path) or os.path.getsize(self.file_path) == 0
        ):
            from openpyxl import Workbook

            wb = Workbook()
            wb.save(self.file_path)
            wb.close()
        self.data = pd.read_excel(self.file_path, sheet_name=None)
        for sheet_name, frame in self.data.items():
            self._validate_columns(sheet_name, frame.columns)

    def _validate_columns(self, sheet_name: str, columns: pd.Index) -> None:
        normalized_headers: set[str] = set()
        normalized_pairs: set[tuple[str, str]] = set()

        for index, column in enumerate(columns, start=1):
            column_name = str(column)
            trimmed = column_name.strip()
            if not trimmed or trimmed.startswith("Unnamed:"):
                raise DataError(f"Empty or None header at column index {index}")

            match = re.match(r"^(?P<base>.+)\.(?P<suffix>[1-9]\d*)$", trimmed)
            if match is not None:
                base = match.group("base").strip()
                base_key = base.casefold()
                if base_key in normalized_headers:
                    raise DataError(
                        f"Duplicate header: '{base}' (sheet '{sheet_name}')"
                    )
                normalized_pairs.add((base_key, trimmed.casefold()))

            header_key = trimmed.casefold()
            if header_key in normalized_headers:
                raise DataError(f"Duplicate header: '{trimmed}' (sheet '{sheet_name}')")
            normalized_headers.add(header_key)

        for base_key, suffixed_key in normalized_pairs:
            if base_key == suffixed_key:
                continue
            if base_key not in normalized_headers:
                continue
            raise DataError(f"Duplicate header detected in sheet '{sheet_name}'")

    def save(self) -> None:
        directory = os.path.dirname(self.file_path) or "."
        temp_file = None
        try:
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=".xlsx", dir=directory
            ) as handle:
                temp_file = handle.name
            os.chmod(temp_file, 0o600)
            with pd.ExcelWriter(temp_file, engine="openpyxl") as writer:
                for sheet_name, frame in self.data.items():
                    frame.to_excel(writer, sheet_name=sheet_name, index=False)
            os.replace(temp_file, self.file_path)
        finally:
            if temp_file and os.path.exists(temp_file):
                os.unlink(temp_file)

    def snapshot(self) -> dict[str, pd.DataFrame]:
        return {name: frame.copy(deep=True) for name, frame in self.data.items()}

    def restore(self, snapshot: Any) -> None:
        self.data = {name: frame.copy(deep=True) for name, frame in snapshot.items()}

    def list_sheets(self) -> list[str]:
        return list(self.data.keys())

    def read_sheet(self, sheet_name: str) -> TableData:
        frame = self.data.get(sheet_name)
        if frame is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")

        row_count = len(frame.index)
        self._check_row_limit(sheet_name, row_count)
        approx_bytes = int(frame.memory_usage(index=True, deep=True).sum())
        self._check_memory_limit(sheet_name, approx_bytes)

        headers = _normalize_headers([str(col) for col in frame.columns])
        rows: list[list[Any]] = []
        for row in frame.itertuples(index=False, name=None):
            row_values: list[Any] = []
            for value in row:
                if pd.isna(value):
                    row_values.append(None)
                else:
                    row_values.append(value)
            rows.append(row_values)
        return TableData(headers=headers, rows=rows)

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        if sheet_name not in self.data:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")
        self.data[sheet_name] = pd.DataFrame(data.rows, columns=pd.Series(data.headers))

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        frame = self.data.get(sheet_name)
        if frame is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in Excel")
        row_data = {col: None for col in frame.columns}
        for idx, col in enumerate(frame.columns):
            if idx < len(row):
                row_data[col] = row[idx]
        self.data[sheet_name] = pd.concat(
            [frame, pd.DataFrame([row_data])], ignore_index=True
        )
        return len(self.data[sheet_name]) + 1

    def create_sheet(self, name: str, headers: list[str]) -> None:
        if name in self.data:
            raise ValueError(f"Sheet '{name}' already exists")
        self.data[name] = pd.DataFrame(columns=pd.Series(headers))

    def drop_sheet(self, name: str) -> None:
        if name not in self.data:
            raise ValueError(f"Sheet '{name}' not found in Excel")
        del self.data[name]

    def get_workbook(self) -> Any:
        raise NotSupportedError(
            f"Backend '{type(self).__name__}' does not expose a workbook object"
        )

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
