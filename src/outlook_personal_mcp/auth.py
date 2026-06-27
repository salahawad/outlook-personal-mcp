from __future__ import annotations

import os
import sys

import msal

from .config import Settings


class AuthRequired(Exception):
    """Raised when interactive device-code login is needed. str() is user-facing."""


class TokenProvider:
    def __init__(self, settings: Settings, app_factory=None):
        self._s = settings
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        factory = app_factory or self._default_app_factory
        self._app = factory(
            settings.client_id,
            authority=settings.authority,
            token_cache=self._cache,
        )

    @staticmethod
    def _default_app_factory(client_id, *, authority, token_cache):
        return msal.PublicClientApplication(
            client_id,
            authority=authority,
            token_cache=token_cache,
        )

    def _load_cache(self):
        path = self._s.token_cache_path
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as fh:
                self._cache.deserialize(fh.read())

    def _save_cache(self, force: bool = False):
        if not force and not self._cache.has_state_changed:
            return
        path = self._s.token_cache_path
        parent = os.path.dirname(path)
        if parent:
            os.makedirs(parent, exist_ok=True)
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(self._cache.serialize())
        os.chmod(path, 0o600)

    def get_token(self) -> str:
        accounts = self._app.get_accounts()
        if accounts:
            res = self._app.acquire_token_silent(self._s.scopes, account=accounts[0])
            if res and "access_token" in res:
                self._save_cache()
                return res["access_token"]
        flow = self._app.initiate_device_flow(scopes=self._s.scopes)
        msg = flow.get("message", "Sign in to continue.")
        user_code = flow.get("user_code", "")
        code_hint = f" (code: {user_code})" if user_code else ""
        raise AuthRequired(
            f"Authentication required{code_hint}. {msg}\n"
            "Run `outlook-personal-mcp login` (or retry this tool after signing in)."
        )

    def login(self) -> None:
        flow = self._app.initiate_device_flow(scopes=self._s.scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Failed to start device flow: {flow.get('error_description', flow)}"
            )
        print(flow["message"], file=sys.stderr, flush=True)
        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(
                f"Login failed: {result.get('error_description', result)}"
            )
        self._save_cache(force=True)
