from typing import Any, Dict, Optional

import pandas as pd

from .base import BaseEngine
from .pandas_executor import PandasExecutor
from .parser import parse_sql
from .result import ExecutionResult


class PandasEngine(BaseEngine):
    """
    PandasEngine loads Excel sheets into DataFrames and executes queries.
    """

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.data = self.load()

    def load(self) -> Dict[str, Any]:
        return pd.read_excel(self.file_path, sheet_name=None)

    def save(self) -> None:
        with pd.ExcelWriter(self.file_path, engine="openpyxl") as writer:
            for sheet_name, frame in self.data.items():
                frame.to_excel(writer, sheet_name=sheet_name, index=False)

    def execute(self, query: str) -> ExecutionResult:
        parsed = parse_sql(query)
        return PandasExecutor(self.data).execute(parsed)

    def execute_with_params(self, query: str, params: Optional[tuple] = None) -> ExecutionResult:
        parsed = parse_sql(query, params)
        return PandasExecutor(self.data).execute(parsed)
