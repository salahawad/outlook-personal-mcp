import pytest
from outlook_personal_mcp.server import build_server
from outlook_personal_mcp.config import Settings

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def mcp(monkeypatch):
    # Inject stub token provider so no MSAL/network during construction.
    from outlook_personal_mcp import server
    monkeypatch.setattr(server, "TokenProvider", lambda *_a, **_k: StubTokens())
    return build_server(Settings(client_id="abc"))

async def test_whoami_registered(mcp):
    tools = await mcp.list_tools()
    names = {t.name for t in tools}
    assert "whoami" in names
    assert "send_mail" in names
    assert "list_events" in names
    assert "create_event" in names
    assert "list_folders" in names
    assert "create_draft" in names
