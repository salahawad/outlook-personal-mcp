from __future__ import annotations
import logging
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from .auth import TokenProvider
from .graph import GraphClient
from .config import Settings
from . import tools

# Third-party loggers that emit full request URLs (including $search terms and
# message/folder IDs) at INFO. For a stdio MCP server those would leak into the
# host's stderr log on every call, so keep them quiet unless OUTLOOK_MCP_DEBUG.
_NOISY_LOGGERS = ("httpx", "httpcore", "msal", "urllib3")


def configure_logging(settings: Settings) -> None:
    """Silence noisy third-party request logging unless debug is enabled."""
    level = logging.INFO if settings.debug else logging.WARNING
    for name in _NOISY_LOGGERS:
        logging.getLogger(name).setLevel(level)


def build_server(settings: Settings) -> FastMCP:
    configure_logging(settings)
    tokens = TokenProvider(settings)
    client = GraphClient(tokens, debug=settings.debug)
    mcp = FastMCP("outlook-personal-mcp")

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get signed-in user profile",
            readOnlyHint=True,
            destructiveHint=False,
            openWorldHint=True,
        )
    )
    async def whoami() -> dict:
        """Return the signed-in user's Microsoft account profile."""
        me = await client.get("/me")
        return {
            "display_name": me.get("displayName"),
            "email": me.get("userPrincipalName") or me.get("mail"),
            "id": me.get("id"),
        }

    tools.register_all(mcp, client, settings)
    return mcp

