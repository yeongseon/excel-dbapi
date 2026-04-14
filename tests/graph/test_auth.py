"""Tests for token provider normalisation and adapters."""

import pytest

from excel_dbapi.engines.graph.auth import (
    AzureIdentityTokenProvider,
    CallbackTokenProvider,
    StaticTokenProvider,
    normalize_token_provider,
)


class TestStaticTokenProvider:
    def test_returns_fixed_token(self):
        tp = StaticTokenProvider("tok-abc")
        assert tp.get_token() == "tok-abc"

    def test_has_get_token_method(self):
        tp = StaticTokenProvider("t")
        assert hasattr(tp, "get_token")
        assert callable(tp.get_token)


class TestCallbackTokenProvider:
    def test_returns_callback_result(self):
        tp = CallbackTokenProvider(lambda: "dynamic-tok")
        assert tp.get_token() == "dynamic-tok"

    def test_calls_each_time(self):
        counter = {"n": 0}

        def factory():
            counter["n"] += 1
            return f"tok-{counter['n']}"

        tp = CallbackTokenProvider(factory)
        assert tp.get_token() == "tok-1"
        assert tp.get_token() == "tok-2"


class TestAzureIdentityTokenProvider:
    def test_wraps_azure_credential(self):
        class FakeCred:
            def get_token(self, scope):
                class T:
                    token = f"azure-{scope}"

                return T()

        tp = AzureIdentityTokenProvider(FakeCred())
        assert tp.get_token() == "azure-https://graph.microsoft.com/.default"


class TestNormalizeTokenProvider:
    def test_string_becomes_static(self):
        tp = normalize_token_provider("my-token")
        assert isinstance(tp, StaticTokenProvider)
        assert tp.get_token() == "my-token"

    def test_callable_becomes_callback(self):
        tp = normalize_token_provider(lambda: "cb-tok")
        assert isinstance(tp, CallbackTokenProvider)
        assert tp.get_token() == "cb-tok"

    def test_token_provider_passthrough(self):
        static = StaticTokenProvider("pass")
        tp = normalize_token_provider(static)
        assert tp is static

    def test_azure_credential_duck_type(self):
        class DuckCred:
            def get_token(self, scope):
                class T:
                    token = "duck"

                return T()

        tp = normalize_token_provider(DuckCred())
        assert isinstance(tp, AzureIdentityTokenProvider)

    def test_azure_credential_varargs_get_token_is_wrapped(self):
        class VarArgsCred:
            def get_token(self, *scopes):
                class T:
                    token = f"varargs-{','.join(scopes)}"

                return T()

        tp = normalize_token_provider(VarArgsCred())
        assert isinstance(tp, AzureIdentityTokenProvider)
        assert tp.get_token() == "varargs-https://graph.microsoft.com/.default"

    def test_none_without_azure_identity_raises(self):
        # This test assumes azure-identity is NOT installed in CI.
        # If it IS installed, the test will still pass (just takes
        # a different branch).
        try:
            tp = normalize_token_provider(None)
            # azure-identity is available — just check it returns something
            assert tp is not None
        except ImportError as exc:
            assert "azure-identity" in str(exc)

    def test_unsupported_type_raises(self):
        with pytest.raises(TypeError, match="Cannot normalise"):
            normalize_token_provider(42)

    def test_custom_zero_arg_get_token_passthrough(self):
        """An object with zero-arg get_token() is treated as a TokenProvider."""

        class CustomTP:
            def get_token(self):
                return "custom-tok"

        cred = CustomTP()
        tp = normalize_token_provider(cred)
        assert tp is cred
        assert tp.get_token() == "custom-tok"
