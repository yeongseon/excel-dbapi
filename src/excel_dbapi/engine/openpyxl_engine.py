from typing import Any, Dict, List
from openpyxl import load_workbook

from .base import BaseEngine
from .executor import execute_query
from .parser import parse_sql


class OpenpyxlEngine(BaseEngine):
    """
    OpenpyxlEngine is responsible for loading, executing, and saving Excel data
    using the openpyxl library. It implements the BaseEngine interface.
    """

    def __init__(self, file_path: str):
        """
        Initialize the OpenpyxlEngine with the given file path.

        Args:
            file_path (str): Path to the Excel (.xlsx) file.
        """
        self.file_path = file_path
        self.data = self.load()

    def load(self) -> Dict[str, Any]:
        """
        Load the Excel workbook into memory.

        Returns:
            Dict[str, Any]: A dictionary mapping sheet names to openpyxl Worksheet objects.
        """
        wb = load_workbook(self.file_path, data_only=True)
        return {sheet: wb[sheet] for sheet in wb.sheetnames}

    def save(self) -> None:
        """
        Save any in-memory changes back to the Excel file.

        Notes:
            - This method reloads the workbook, updates each worksheet's cells
              with the current in-memory data, and saves it back to the file.
            - Intended for future use when INSERT, UPDATE, DELETE operations are supported.
        """
        wb = load_workbook(self.file_path)
        for sheet_name, sheet in self.data.items():
            ws = wb[sheet_name]
            for row_idx, row in enumerate(sheet.iter_rows(), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    ws.cell(row=row_idx, column=col_idx, value=cell.value)
        wb.save(self.file_path)

    def execute(self, query: str) -> List[Dict[str, Any]]:
        """
        Execute a SQL-like query on the loaded Excel data.

        Args:
            query (str): A SQL-like query string (e.g., "SELECT * FROM Sheet1 WHERE id = '1'").

        Returns:
            List[Dict[str, Any]]: Query results as a list of dictionaries, where each dictionary
                                  represents a row with column names as keys.
        """
        parsed = parse_sql(query)
        return execute_query(parsed, self.data)
