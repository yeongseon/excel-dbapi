from abc import ABC, abstractmethod
from typing import Any, Dict, List


class BaseEngine(ABC):
    @abstractmethod
    def load(self) -> Dict[str, Any]:
        """
        Load data from the Excel file.
        """
        pass

    @abstractmethod
    def execute(self, query: str) -> List[Dict[str, Any]]:
        """
        Execute a query against the loaded data.
        """
        pass
