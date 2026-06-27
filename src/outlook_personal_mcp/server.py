from __future__ import annotations
from mcp.server.fastmcp import FastMCP
from .auth import TokenProvider
from .graph import GraphClient
from .config import Settings
from . import tools


def build_server(settings: Settings) -> FastMCP:
    tokens = TokenProvider(settings)
    client = GraphClient(tokens, debug=settings.debug)
    mcp = FastMCP("outlook-personal-mcp")

    @mcp.tool()
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

