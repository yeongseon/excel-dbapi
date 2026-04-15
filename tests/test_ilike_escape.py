from typing import Any

from excel_dbapi.executor import SharedExecutor
from excel_dbapi.engines.base import TableData, WorkbookBackend
from excel_dbapi.parser import parse_sql


class _StubBackend(WorkbookBackend):
    @property
    def readonly(self) -> bool:
        return False

    @property
    def supports_transactions(self) -> bool:
        return True

    def __init__(self) -> None:
        super().__init__("stub.xlsx")

    def load(self) -> None:
        return None

    def save(self) -> None:
        return None

    def snapshot(self) -> object:
        return object()

    def restore(self, snapshot: object) -> None:
        return None

    def list_sheets(self) -> list[str]:
        return []

    def read_sheet(self, sheet_name: str) -> TableData:
        raise NotImplementedError

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        raise NotImplementedError

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        raise NotImplementedError

    def create_sheet(self, name: str, headers: list[str]) -> None:
        raise NotImplementedError

    def drop_sheet(self, name: str) -> None:
        raise NotImplementedError


def _executor() -> SharedExecutor:
    return SharedExecutor(_StubBackend())


def test_parse_ilike_and_escape_clause() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE name ILIKE 'a!_%' ESCAPE '!'")
    assert parsed["where"]["conditions"][0] == {
        "column": "name",
        "operator": "ILIKE",
        "value": "a!_%",
        "escape": "!",
    }


def test_parse_not_ilike() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE name NOT ILIKE 'a%'")
    assert parsed["where"]["conditions"][0] == {
        "column": "name",
        "operator": "NOT ILIKE",
        "value": "a%",
    }


def test_ilike_case_insensitive_and_wildcards() -> None:
    executor = _executor()
    assert executor._evaluate_condition(
        {"name": "ALPHA"},
        {"column": "name", "operator": "ILIKE", "value": "%ph%"},
    )
    assert executor._evaluate_condition(
        {"name": "Alice"},
        {"column": "name", "operator": "ILIKE", "value": "a%"},
    )


def test_not_ilike() -> None:
    executor = _executor()
    assert executor._evaluate_condition(
        {"name": "Bob"},
        {"column": "name", "operator": "NOT ILIKE", "value": "a%"},
    )
    assert not executor._evaluate_condition(
        {"name": "Alice"},
        {"column": "name", "operator": "NOT ILIKE", "value": "a%"},
    )


def test_like_and_ilike_escape_clause() -> None:
    executor = _executor()
    assert executor._evaluate_condition(
        {"tag": "100% pure"},
        {"column": "tag", "operator": "LIKE", "value": "100!% pure", "escape": "!"},
    )
    assert executor._evaluate_condition(
        {"name": "A_test"},
        {"column": "name", "operator": "ILIKE", "value": "a!_%", "escape": "!"},
    )


def test_escape_percent_and_underscore() -> None:
    executor = _executor()
    assert executor._evaluate_condition(
        {"tag": "50%_done"},
        {"column": "tag", "operator": "LIKE", "value": "%!%%", "escape": "!"},
    )
    assert executor._evaluate_condition(
        {"tag": "a_b"},
        {"column": "tag", "operator": "LIKE", "value": "a!_b", "escape": "!"},
    )
