from typing import Callable

from ..exceptions import CapabilityError
from .base import WorkbookBackend

_REGISTRY: dict[str, Callable[..., type[WorkbookBackend]]] = {}
_SCHEME_MAP: dict[str, str] = {}


def register_engine(
    name: str,
    factory: Callable[..., type[WorkbookBackend]],
    schemes: tuple[str, ...] = (),
) -> None:
    key = name.lower()
    _REGISTRY[key] = factory
    for scheme in schemes:
        _SCHEME_MAP[scheme.lower()] = key


def get_engine(name: str) -> type[WorkbookBackend]:
    key = name.lower()
    try:
        loader = _REGISTRY[key]
    except KeyError as exc:
        available = ", ".join(sorted(_REGISTRY.keys()))
        raise CapabilityError(f"Unsupported engine: {name}. Available engines: {available}") from exc
    return loader()


def resolve_engine_from_dsn(dsn: str) -> str | None:
    if "://" not in dsn:
        return None
    scheme = dsn.split("://", 1)[0].lower()
    return _SCHEME_MAP.get(scheme)


def _load_openpyxl() -> type[WorkbookBackend]:
    from .openpyxl.backend import OpenpyxlBackend

    return OpenpyxlBackend


def _load_pandas() -> type[WorkbookBackend]:
    try:
        from .pandas.backend import PandasBackend
    except ModuleNotFoundError as exc:
        if exc.name not in {"pandas", "openpyxl"}:
            raise
        raise CapabilityError(
            "The pandas engine requires the 'pandas' package: "
            "pip install 'excel-dbapi[pandas]'"
        ) from exc
    return PandasBackend


register_engine("openpyxl", _load_openpyxl)
register_engine("pandas", _load_pandas)


def _load_graph() -> type[WorkbookBackend]:
    try:
        from .graph.backend import GraphBackend
    except ModuleNotFoundError as exc:
        if exc.name not in {"httpx"}:
            raise
        raise CapabilityError(
            "The graph engine requires the 'httpx' package: "
            "pip install 'excel-dbapi[graph]'"
        ) from exc
    return GraphBackend


register_engine("graph", _load_graph, schemes=("msgraph", "sharepoint", "onedrive"))
