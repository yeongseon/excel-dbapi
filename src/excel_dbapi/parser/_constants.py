from typing import Any, SupportsIndex

from ..exceptions import SqlSemanticError


class _QuotedString(str):
    pass


class _QuotedIdentifier(str):
    pass


class _OrderByClause(list[dict[str, Any]]):
    def __getitem__(self, index: SupportsIndex | slice | str) -> Any:
        if isinstance(index, str):
            if len(self) != 1:
                raise TypeError("list indices must be integers or slices, not str")
            return super().__getitem__(0)[index]
        return super().__getitem__(index)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, dict):
            return len(self) == 1 and super().__getitem__(0) == other
        return super().__eq__(other)


def _is_placeholder(value: Any) -> bool:
    return value == "?" and not isinstance(value, _QuotedString)


_COLUMN_TYPE_ALIASES = {
    "INT": "INTEGER",
    "FLOAT": "REAL",
}

_SUPPORTED_COLUMN_TYPES = {"TEXT", "INTEGER", "REAL", "BOOLEAN", "DATE", "DATETIME"}


def _normalize_column_type(type_name: str, *, context: str) -> str:
    normalized = _COLUMN_TYPE_ALIASES.get(type_name.upper(), type_name.upper())
    if normalized not in _SUPPORTED_COLUMN_TYPES:
        raise SqlSemanticError(f"Unsupported {context} column type: {type_name}")
    return normalized


_AGGREGATE_FUNCTIONS = frozenset({"COUNT", "SUM", "AVG", "MIN", "MAX"})
_WINDOW_FUNCTIONS = frozenset({"ROW_NUMBER", "RANK", "DENSE_RANK"})
_SCALAR_FUNCTION_NAMES = frozenset(
    {
        "COALESCE",
        "NULLIF",
        "UPPER",
        "LOWER",
        "TRIM",
        "LENGTH",
        "SUBSTR",
        "SUBSTRING",
        "ABS",
        "ROUND",
        "REPLACE",
        "CONCAT",
        "YEAR",
        "MONTH",
        "DAY",
    }
)
_IDENTIFIER_PATTERN = r"[A-Za-z_\u0080-\uffff][A-Za-z0-9_\u0080-\uffff]*"
_QUALIFIED_IDENTIFIER_PATTERN = rf"{_IDENTIFIER_PATTERN}\.{_IDENTIFIER_PATTERN}"
