from typing import Any, Dict

from excel_dbapi.engine.base import BaseEngine
from excel_dbapi.engine.result import ExecutionResult


class DummyEngine(BaseEngine):
    def __init__(self) -> None:
        self.called = False

    def load(self) -> Dict[str, Any]:
        return {}

    def save(self) -> None:
        self.called = True

    def snapshot(self) -> Any:
        return {"state": "ok"}

    def restore(self, snapshot: Any) -> None:
        self.called = True

    def execute(self, query: str) -> ExecutionResult:
        return ExecutionResult(action="SELECT", rows=[], description=[], rowcount=0, lastrowid=None)


def test_execute_with_params_delegates() -> None:
    engine = DummyEngine()
    result = engine.execute_with_params("SELECT 1", (1,))
    assert result.rowcount == 0


def test_base_methods_coverage() -> None:
    engine = DummyEngine()
    BaseEngine.load(engine)
    BaseEngine.save(engine)
    BaseEngine.snapshot(engine)
    BaseEngine.restore(engine, {})
