from __future__ import annotations
import os
from dataclasses import dataclass, field

DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"
# offline_access is added by MSAL automatically; do not list it here.
SCOPES = ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "User.Read"]
DEFAULT_MAX_FILE_BYTES = 3 * 1024 * 1024


def _default_cache_path() -> str:
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, "outlook-personal-mcp", "token_cache.bin")


def _default_file_root() -> str:
    base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    return os.path.join(base, "outlook-personal-mcp", "files")


def _truthy(v: str | None) -> bool:
    return str(v).strip().lower() in {"1", "true", "yes", "on"}


def _positive_int(v: str | None, *, default: int, name: str) -> int:
    if v is None or not str(v).strip():
        return default
    try:
        value = int(str(v).strip())
    except ValueError as exc:
        raise ValueError(f"{name} must be a positive integer") from exc
    if value <= 0:
        raise ValueError(f"{name} must be a positive integer")
    return value


@dataclass(frozen=True)
class Settings:
    client_id: str
    authority: str = DEFAULT_AUTHORITY
    scopes: list[str] = field(default_factory=lambda: list(SCOPES))
    token_cache_path: str = field(default_factory=_default_cache_path)
    file_root: str = field(default_factory=_default_file_root)
    max_file_bytes: int = DEFAULT_MAX_FILE_BYTES
    allow_permanent_delete: bool = False
    debug: bool = False


def load_settings(environ: dict | None = None) -> Settings:
    env = os.environ if environ is None else environ
    client_id = env.get("OUTLOOK_MCP_CLIENT_ID", "").strip()
    if not client_id:
        raise ValueError("OUTLOOK_MCP_CLIENT_ID is required")
    cache = env.get("OUTLOOK_MCP_TOKEN_CACHE")
    file_root = env.get("OUTLOOK_MCP_FILE_ROOT")
    return Settings(
        client_id=client_id,
        authority=env.get("OUTLOOK_MCP_AUTHORITY", DEFAULT_AUTHORITY),
        token_cache_path=os.path.expanduser(cache) if cache else _default_cache_path(),
        file_root=os.path.abspath(os.path.expanduser(file_root))
        if file_root
        else _default_file_root(),
        max_file_bytes=_positive_int(
            env.get("OUTLOOK_MCP_MAX_FILE_BYTES"),
            default=DEFAULT_MAX_FILE_BYTES,
            name="OUTLOOK_MCP_MAX_FILE_BYTES",
        ),
        allow_permanent_delete=_truthy(env.get("OUTLOOK_MCP_ALLOW_PERMANENT_DELETE")),
        debug=_truthy(env.get("OUTLOOK_MCP_DEBUG")),
    )
