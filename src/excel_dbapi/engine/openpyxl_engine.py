from typing import Any, Dict, List
from openpyxl import load_workbook

from .base import BaseEngine
from .executor import execute_query
from .parser import parse_sql


class OpenpyxlEngine(BaseEngine):
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.data = self.load()

    def load(self) -> Dict[str, Any]:
        wb = load_workbook(self.file_path, data_only=True)
        return {sheet: wb[sheet] for sheet in wb.sheetnames}

    def save(self) -> None:
        wb = load_workbook(self.file_path)
        for sheet_name, sheet in self.data.items():
            ws = wb[sheet_name]
            for row_idx, row in enumerate(sheet.iter_rows(), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    ws.cell(row=row_idx, column=col_idx, value=cell.value)
        wb.save(self.file_path)

    def execute(self, query: str) -> List[Dict[str, Any]]:
        parsed = parse_sql(query)
        return execute_query(parsed, self.data)
