import pytest
from outlook_personal_mcp.config import load_settings

def test_load_requires_client_id(monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_CLIENT_ID", raising=False)
    with pytest.raises(ValueError, match="OUTLOOK_MCP_CLIENT_ID"):
        load_settings(environ={})

def test_load_defaults(monkeypatch):
    s = load_settings(environ={"OUTLOOK_MCP_CLIENT_ID": "abc"})
    assert s.client_id == "abc"
    assert s.authority == "https://login.microsoftonline.com/consumers"
    assert s.scopes == ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "User.Read"]
    assert s.allow_permanent_delete is False
    assert s.token_cache_path.endswith("token_cache.bin")

def test_allow_permanent_delete_truthy():
    s = load_settings(environ={"OUTLOOK_MCP_CLIENT_ID": "abc",
                               "OUTLOOK_MCP_ALLOW_PERMANENT_DELETE": "true"})
    assert s.allow_permanent_delete is True
