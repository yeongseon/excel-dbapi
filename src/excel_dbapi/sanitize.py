"""Formula injection defense for Excel cell values.

OWASP recommends escaping cell values that begin with characters
interpreted as formulas by spreadsheet applications.  When
``sanitize_formulas`` is enabled (the default), string values starting
with ``=``, ``+``, ``-``, ``@``, ``\\t``, or ``\\r`` are prefixed with a
leading single-quote (``'``) so that Excel treats them as literal text.

References
----------
- OWASP CSV Injection: https://owasp.org/www-community/attacks/CSV_Injection
"""

from typing import Any, List, Sequence

# Characters that trigger formula interpretation in Excel / Google Sheets.
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def sanitize_cell_value(value: Any) -> Any:
    """Sanitize a single cell value against formula injection.

    If *value* is a string starting with a known formula prefix, a leading
    single-quote is prepended.  Non-string values are returned unchanged.

    Args:
        value: The cell value to sanitize.

    Returns:
        The sanitized value.
    """
    if isinstance(value, str) and value and value[0] in _FORMULA_PREFIXES:
        return "'" + value
    return value


def sanitize_row(row: Sequence[Any]) -> List[Any]:
    """Sanitize every value in a row.

    Args:
        row: A sequence of cell values (e.g. a list or tuple).

    Returns:
        A new list with each element sanitized.
    """
    return [sanitize_cell_value(v) for v in row]


__all__ = [
    "sanitize_cell_value",
    "sanitize_row",
]
