from datetime import date, datetime, time
import re
from typing import Any, Callable, Protocol

from ..exceptions import SqlSemanticError

_READONLY_ACTIONS = frozenset({"INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "ALTER"})



class _SupportsOrder(Protocol):
    def __lt__(self, other: Any, /) -> bool: ...


def _build_like_regex(pattern: str, escape_char: str | None) -> str:
    parts: list[str] = ["^"]
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if escape_char is not None and char == escape_char:
            index += 1
            if index >= len(pattern):
                raise SqlSemanticError("Invalid LIKE pattern: trailing ESCAPE character")
            parts.append(re.escape(pattern[index]))
        elif char == "%":
            parts.append(".*")
        elif char == "_":
            parts.append(".")
        else:
            parts.append(re.escape(char))
        index += 1

    parts.append("$")
    return "".join(parts)


ScalarFunctionHandler = Callable[[list[Any]], Any]
ScalarFunctionSpec = tuple[int, int | None, ScalarFunctionHandler]


def _coalesce(args: list[Any]) -> Any:
    for value in args:
        if value is not None:
            return value
    return None


def _nullif(args: list[Any]) -> Any:
    return None if args[0] == args[1] else args[0]


def _upper(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else str(value).upper()


def _lower(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else str(value).lower()


def _trim(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else str(value).strip()


def _length(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else len(str(value))


def _to_int_like(value: Any) -> int:
    if isinstance(value, bool):
        raise SqlSemanticError("expected numeric value")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise SqlSemanticError("expected numeric value")
        return int(float(text))
    raise SqlSemanticError("expected numeric value")


def _substr(args: list[Any]) -> Any:
    text_value = args[0]
    start_value = args[1]
    if text_value is None or start_value is None:
        return None

    text = str(text_value)
    start = _to_int_like(start_value)
    if start > 0:
        start_index = start - 1
    elif start < 0:
        start_index = len(text) + start
    else:
        start_index = 0

    if len(args) < 3:
        return text[start_index:]

    length_value = args[2]
    if length_value is None:
        return None
    length = _to_int_like(length_value)
    if length <= 0:
        return ""
    return text[start_index : start_index + length]


def _concat(args: list[Any]) -> str:
    return "".join(str(value) for value in args if value is not None)


def _abs(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    if isinstance(value, bool):
        raise SqlSemanticError("expected numeric value")
    if isinstance(value, (int, float)):
        return abs(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise SqlSemanticError("expected numeric value")
        return abs(float(text))
    raise SqlSemanticError("expected numeric value")


def _round(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    if isinstance(value, bool):
        raise SqlSemanticError("expected numeric value")
    if isinstance(value, (int, float)):
        numeric = float(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            raise SqlSemanticError("expected numeric value")
        numeric = float(text)
    else:
        raise SqlSemanticError("expected numeric value")

    if len(args) < 2 or args[1] is None:
        return round(numeric)
    precision = _to_int_like(args[1])
    return round(numeric, precision)


def _replace(args: list[Any]) -> Any:
    source = args[0]
    if source is None:
        return None
    old = args[1]
    if old is None:
        return str(source)
    new = args[2]
    return str(source).replace(str(old), "" if new is None else str(new))


def _date_value(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None) if value.tzinfo is not None else value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, str):
        normalized = value.strip()
        if not normalized:
            raise SqlSemanticError("expected date value")
        if normalized.endswith("Z"):
            normalized = normalized[:-1] + "+00:00"
        try:
            parsed_datetime = datetime.fromisoformat(normalized)
        except ValueError:
            parsed_datetime = None
        if parsed_datetime is not None:
            return (
                parsed_datetime.replace(tzinfo=None)
                if parsed_datetime.tzinfo is not None
                else parsed_datetime
            )
        try:
            parsed_date = date.fromisoformat(value.strip())
        except ValueError as exc:
            raise SqlSemanticError("expected date value") from exc
        return datetime.combine(parsed_date, time.min)
    raise SqlSemanticError("expected date value")


def _year(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    return _date_value(value).year


def _month(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    return _date_value(value).month


def _day(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    return _date_value(value).day


_SCALAR_FUNCTIONS: dict[str, ScalarFunctionSpec] = {
    "COALESCE": (1, None, _coalesce),
    "NULLIF": (2, 2, _nullif),
    "UPPER": (1, 1, _upper),
    "LOWER": (1, 1, _lower),
    "TRIM": (1, 1, _trim),
    "LENGTH": (1, 1, _length),
    "SUBSTR": (2, 3, _substr),
    "SUBSTRING": (2, 3, _substr),
    "ABS": (1, 1, _abs),
    "ROUND": (1, 2, _round),
    "REPLACE": (3, 3, _replace),
    "CONCAT": (1, None, _concat),
    "YEAR": (1, 1, _year),
    "MONTH": (1, 1, _month),
    "DAY": (1, 1, _day),
}


def _tv_and(a: bool | None, b: bool | None) -> bool | None:
    """SQL three-valued AND."""
    if a is False or b is False:
        return False
    if a is None or b is None:
        return None
    return True


def _tv_or(a: bool | None, b: bool | None) -> bool | None:
    """SQL three-valued OR."""
    if a is True or b is True:
        return True
    if a is None or b is None:
        return None
    return False
