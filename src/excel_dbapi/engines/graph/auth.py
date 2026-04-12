"""Token providers for Microsoft Graph API authentication."""

from __future__ import annotations

from typing import Any, Callable, Protocol, cast


class TokenProvider(Protocol):
    """Protocol for objects that supply a bearer token string."""

    def get_token(self) -> str:
        """Return a valid access token (may refresh internally)."""
        ...


class StaticTokenProvider:
    """Token provider backed by a fixed string — useful for tests or short-lived scripts."""

    def __init__(self, token: str) -> None:
        self._token = token

    def get_token(self) -> str:
        return self._token


class CallbackTokenProvider:
    """Token provider backed by a user-supplied callable."""

    def __init__(self, callback: Callable[[], str]) -> None:
        self._callback = callback

    def get_token(self) -> str:
        return self._callback()


class AzureIdentityTokenProvider:
    """Adapter for ``azure.identity`` credential objects.

    Wraps any credential exposing ``get_token(scopes)`` such as
    ``DefaultAzureCredential``.
    """

    _GRAPH_SCOPE = "https://graph.microsoft.com/.default"

    def __init__(self, credential: Any) -> None:
        self._credential = credential

    def get_token(self) -> str:
        token = self._credential.get_token(self._GRAPH_SCOPE)
        return cast(str, token.token)


def _has_get_token_with_args(obj: Any) -> bool:
    """Return True if *obj* has a ``get_token`` that requires positional args.

    This distinguishes azure-identity style ``get_token(scope)`` from our own
    zero-arg ``get_token()``.
    """
    method = getattr(obj, "get_token", None)
    if method is None:
        return False
    import inspect

    try:
        sig = inspect.signature(method)
    except (ValueError, TypeError):
        return False
    required = [
        p
        for p in sig.parameters.values()
        if p.name != "self"
        and p.default is inspect.Parameter.empty
        and p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)
    ]
    return len(required) >= 1


def normalize_token_provider(credential: Any) -> TokenProvider:
    """Coerce various credential shapes into a ``TokenProvider``.

    Accepted inputs:
    - ``str`` — wrapped in ``StaticTokenProvider``
    - Concrete provider (``StaticTokenProvider``, ``CallbackTokenProvider``,
      ``AzureIdentityTokenProvider``) — returned as-is
    - Object with ``get_token(scope)`` (azure-identity style) →
      ``AzureIdentityTokenProvider``
    - Zero-arg callable → ``CallbackTokenProvider``
    - ``None`` — attempt ``DefaultAzureCredential()`` if azure-identity
      is installed; otherwise raise.

    Raises:
        TypeError: If the credential cannot be normalised.
        ImportError: If ``None`` is passed and ``azure-identity`` is not
            installed.
    """
    # 1. Plain string → static token
    if isinstance(credential, str):
        return StaticTokenProvider(credential)

    # 2. Already one of our concrete providers → pass through
    if isinstance(
        credential,
        (StaticTokenProvider, CallbackTokenProvider, AzureIdentityTokenProvider),
    ):
        return credential

    # 3. Azure-identity style credential (get_token requires a scope arg)
    if _has_get_token_with_args(credential):
        return AzureIdentityTokenProvider(credential)

    # 4. Object with zero-arg get_token() (custom TokenProvider)
    if hasattr(credential, "get_token") and callable(credential.get_token):
        return cast(TokenProvider, credential)

    # 5. Plain callable → callback provider
    if callable(credential):
        return CallbackTokenProvider(cast(Callable[[], str], credential))

    # 6. None → try DefaultAzureCredential
    if credential is None:
        try:
            from azure.identity import DefaultAzureCredential
        except ImportError as exc:
            raise ImportError(
                "No credential provided and azure-identity is not installed. "
                "Install with: pip install excel-dbapi[graph-azure]"
            ) from exc
        return AzureIdentityTokenProvider(DefaultAzureCredential())

    raise TypeError(
        f"Cannot normalise {type(credential).__name__!r} to a TokenProvider"
    )
