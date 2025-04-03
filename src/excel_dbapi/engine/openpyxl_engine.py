from typing import Any, Dict, List

from openpyxl import load_workbook

from .base import BaseEngine


class OpenpyxlEngine(BaseEngine):
    def __init__(self, file_path: str):
        """
        Initialize OpenpyxlEngine with the given file path.
        """
        self.file_path = file_path
        self.data = self.load()

    def load(self) -> Dict[str, Any]:
        """
        Load all sheets using openpyxl.
        """
        wb = load_workbook(self.file_path, data_only=True)
        return {sheet: wb[sheet] for sheet in wb.sheetnames}

    def save(self) -> None:
        """
        Save is not implemented for OpenpyxlEngine.
        """
        raise NotImplementedError(
            "Save operation is not implemented for OpenpyxlEngine."
        )

    def execute(self, query: str) -> List[Dict[str, Any]]:
        """
        Example execution: return all records from the first sheet.
        """
        sheet = list(self.data.keys())[0]
        ws = self.data[sheet]
        rows = list(ws.iter_rows(values_only=True))
        headers = rows[0]
        return [dict(zip(headers, row)) for row in rows[1:]]
