from typing import Any, Dict

import pandas as pd
from pandas import DataFrame

from .base import BaseEngine


class PandasEngine(BaseEngine):
    """
    PandasEngine uses pandas with openpyxl backend to load and query Excel files.
    """

    def __init__(self, file_path: str):
        """
        Initialize PandasEngine with the given file path.
        """
        self.file_path = file_path
        self.data = self.load()

    def load(self) -> Dict[str, DataFrame]:
        """
        Load all sheets as DataFrames using pandas with openpyxl engine.
        """
        try:
            data = pd.read_excel(self.file_path, sheet_name=None, engine="openpyxl")
            return data
        except Exception as e:
            raise IOError(f"Failed to load Excel file: {self.file_path}") from e

    def save(self, file_path: str) -> None:
        """
        Save the current data back to an Excel file.
        """
        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                for sheet_name, df in self.data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            raise IOError(f"Failed to save Excel file: {file_path}") from e

    def execute(self, query: str) -> list[dict[str, Any]]:
        """
        Execute a query and return the result.

        Currently, only SELECT * FROM [Sheet$] is supported as an example.
        """
        sheet = list(self.data.keys())[0]
        return self.data[sheet].to_dict(orient="records")
