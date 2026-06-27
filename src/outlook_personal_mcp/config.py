from __future__ import annotations
import os
from dataclasses import dataclass, field

DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"
# offline_access is added by MSAL automatically; do not list it here.
SCOPES = ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "User.Read"]


def _default_cache_path() -> str:
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, "outlook-personal-mcp", "token_cache.bin")


def _truthy(v: str | None) -> bool:
    return str(v).strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    client_id: str
    authority: str = DEFAULT_AUTHORITY
    scopes: list[str] = field(default_factory=lambda: list(SCOPES))
    token_cache_path: str = field(default_factory=_default_cache_path)
    allow_permanent_delete: bool = False
    debug: bool = False


def load_settings(environ: dict | None = None) -> Settings:
    env = os.environ if environ is None else environ
    client_id = env.get("OUTLOOK_MCP_CLIENT_ID", "").strip()
    if not client_id:
        raise ValueError("OUTLOOK_MCP_CLIENT_ID is required")
    cache = env.get("OUTLOOK_MCP_TOKEN_CACHE")
    return Settings(
        client_id=client_id,
        authority=env.get("OUTLOOK_MCP_AUTHORITY", DEFAULT_AUTHORITY),
        token_cache_path=os.path.expanduser(cache) if cache else _default_cache_path(),
        allow_permanent_delete=_truthy(env.get("OUTLOOK_MCP_ALLOW_PERMANENT_DELETE")),
        debug=_truthy(env.get("OUTLOOK_MCP_DEBUG")),
    )
