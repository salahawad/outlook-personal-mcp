import os
import stat
import pytest
from outlook_personal_mcp.auth import TokenProvider, AuthRequired
from outlook_personal_mcp.config import Settings


class FakeMsalApp:
    def __init__(self, *, silent=None, flow=None, by_flow=None, accounts=None):
        self._silent, self._flow, self._by_flow = silent, flow, by_flow
        self._accounts = accounts or []

    def get_accounts(self):
        return self._accounts

    def acquire_token_silent(self, scopes, account):
        return self._silent

    def initiate_device_flow(self, scopes):
        return self._flow

    def acquire_token_by_device_flow(self, flow):
        return self._by_flow


def _settings(tmp_path):
    return Settings(client_id="abc", token_cache_path=str(tmp_path / "cache.bin"))


def test_silent_token_returns_access_token(tmp_path):
    app = FakeMsalApp(silent={"access_token": "TOK"}, accounts=[{"username": "u"}])
    tp = TokenProvider(_settings(tmp_path), app_factory=lambda *_a, **_k: app)
    assert tp.get_token() == "TOK"


def test_no_account_raises_auth_required(tmp_path):
    app = FakeMsalApp(
        silent=None,
        accounts=[],
        flow={
            "verification_uri": "https://aka.ms/devicelogin",
            "user_code": "ABCD-1234",
            "message": "Go authenticate",
        },
    )
    tp = TokenProvider(_settings(tmp_path), app_factory=lambda *_a, **_k: app)
    with pytest.raises(AuthRequired) as ei:
        tp.get_token()
    assert "ABCD-1234" in str(ei.value)


def test_login_persists_cache(tmp_path):
    app = FakeMsalApp(
        flow={"user_code": "X", "message": "m"},
        by_flow={"access_token": "TOK", "token_type": "Bearer"},
    )
    s = _settings(tmp_path)
    tp = TokenProvider(s, app_factory=lambda *_a, **_k: app)
    tp.login()  # should not raise
    assert os.path.exists(s.token_cache_path)
    assert stat.S_IMODE(os.stat(s.token_cache_path).st_mode) == 0o600
