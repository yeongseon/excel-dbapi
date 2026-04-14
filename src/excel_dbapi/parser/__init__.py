from typing import Any, Dict, Optional

from .tokenizer import _count_unquoted_placeholders, _tokenize
from .compound import _parse_compound, _parse_with_query
from .dml import _parse_delete, _parse_insert, _parse_update
from .ddl import _parse_alter, _parse_create, _parse_drop
from .select import _parse_select


def parse_sql(query: str, params: Optional[tuple[Any, ...]] = None) -> Dict[str, Any]:
    if params and _count_unquoted_placeholders(query) == 0:
        raise ValueError("Too many parameters for placeholders")

    tokens = _tokenize(query.strip())
    if not tokens:
        raise ValueError(f"Invalid SQL query format: {query}")
    action = tokens[0].upper()
    parsed: Dict[str, Any]
    if action == "WITH":
        parsed = _parse_with_query(query, params)
    elif action == "SELECT" or tokens[0] == "(":
        compound_parsed = _parse_compound(query, params)
        if compound_parsed is None:
            if tokens[0] == "(":
                raise ValueError(f"Unsupported SQL action: {tokens[0]}")
            parsed = _parse_select(query, params)
        else:
            parsed = compound_parsed
    elif action == "INSERT":
        parsed = _parse_insert(query, params)
    elif action == "CREATE":
        parsed = _parse_create(query)
    elif action == "DROP":
        parsed = _parse_drop(query)
    elif action == "UPDATE":
        parsed = _parse_update(query, params)
    elif action == "DELETE":
        parsed = _parse_delete(query, params)
    elif action == "ALTER":
        parsed = _parse_alter(query)
    else:
        raise ValueError(f"Unsupported SQL action: {action}")

    parsed["params"] = params
    return parsed


__all__ = [
    "parse_sql",
    # _constants
    "_OrderByClause",
    "_QuotedIdentifier",
    "_QuotedString",
    "_is_placeholder",
    # tokenizer
    "_count_unquoted_placeholders",
    "_find_matching_parenthesis",
    "_find_top_level_keyword_index",
    "_is_double_quoted_token",
    "_is_identifier_or_quoted",
    "_is_quoted_token",
    "_is_qualified_identifier_or_quoted",
    "_is_single_quoted_token",
    "_parse_column_identifier",
    "_parse_value",
    "_split_qualified_identifier",
    "_split_csv",
    "_tokenize",
    "_tokenize_expression",
    # expressions
    "_annotate_column_tables",
    "_apply_bound_values_to_condition",
    "_bind_expression_values",
    "_bind_where_conditions",
    "_collect_case_tokens_until",
    "_collect_qualified_references_from_expression",
    "_expression_to_sql_for_order_by",
    "_expression_values_to_bind",
    "_normalize_aggregate_expressions",
    "_parse_case_expression",
    "_parse_case_expression_tokens",
    "_parse_column_expression",
    "_parse_window_spec_tokens",
    "_values_to_bind_from_condition",
    "_where_values_to_bind",
    # where
    "_collect_qualified_references_from_where",
    "_is_subquery_condition",
    "_parse_where_expression",
    "_where_to_sql_for_order_by",
    # select
    "_bind_params",
    "_collect_qualified_references_from_query",
    "_find_clause_positions",
    "_parse_columns",
    "_parse_order_by_clause_tokens",
    "_parse_order_by_item_tokens",
    "_parse_select",
    "_query_source_references",
    "_validate_join_column_reference",
    "_validate_join_on_condition_node",
    "_validate_join_where_node",
    # dml
    "_parse_delete",
    "_parse_insert",
    "_parse_on_conflict_clause",
    "_parse_update",
    # ddl
    "_parse_alter",
    "_parse_create",
    "_parse_drop",
    # compound
    "_parse_compound",
    "_parse_with_query",
    "_query_references_name",
]

# --- Backwards-compatibility re-exports for tests and executor ---

from ._constants import (  # noqa: F401, E402
    _OrderByClause,
    _QuotedIdentifier,
    _QuotedString,
    _is_placeholder,
)
from .tokenizer import (  # noqa: F401, E402
    _find_matching_parenthesis,
    _find_top_level_keyword_index,
    _is_double_quoted_token,
    _is_identifier_or_quoted,
    _is_quoted_token,
    _is_qualified_identifier_or_quoted,
    _is_single_quoted_token,
    _parse_column_identifier,
    _parse_value,
    _split_csv,
    _split_qualified_identifier,
    _tokenize_expression,
)
from .expressions import (  # noqa: F401, E402
    _annotate_column_tables,
    _bind_expression_values,
    _bind_where_conditions,
    _collect_case_tokens_until,
    _collect_qualified_references_from_expression,
    _expression_to_sql_for_order_by,
    _expression_values_to_bind,
    _normalize_aggregate_expressions,
    _parse_case_expression,
    _parse_case_expression_tokens,
    _parse_column_expression,
    _parse_window_spec_tokens,
    _values_to_bind_from_condition,
    _apply_bound_values_to_condition,
    _where_values_to_bind,
)
from .where import (  # noqa: F401, E402
    _collect_qualified_references_from_where,
    _is_subquery_condition,
    _parse_where_expression,
    _where_to_sql_for_order_by,
)
from .select import (  # noqa: F401, E402
    _bind_params,
    _collect_qualified_references_from_query,
    _find_clause_positions,
    _parse_columns,
    _parse_order_by_clause_tokens,
    _parse_order_by_item_tokens,
    _query_source_references,
    _validate_join_column_reference,
    _validate_join_on_condition_node,
    _validate_join_where_node,
)
from .dml import _parse_on_conflict_clause  # noqa: F401, E402
from .compound import _query_references_name  # noqa: F401, E402
