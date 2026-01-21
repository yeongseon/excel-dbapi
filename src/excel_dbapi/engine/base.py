from abc import ABC, abstractmethod
from typing import Any, Dict, Optional

from .result import ExecutionResult


class BaseEngine(ABC):
    @abstractmethod
    def load(self) -> Dict[str, Any]:
        """
        Load data from the Excel file.
        """
        pass

    @abstractmethod
    def save(self) -> None:
        """Persist in-memory changes to disk."""
        pass

    @abstractmethod
    def execute(self, query: str) -> ExecutionResult:
        """
        Execute a query against the loaded data.
        """
        pass

    def execute_with_params(self, query: str, params: Optional[tuple] = None) -> ExecutionResult:
        """Execute a query with optional parameters."""
        return self.execute(query)
