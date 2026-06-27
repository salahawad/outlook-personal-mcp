import pytest
from outlook_personal_mcp.config import DEFAULT_MAX_FILE_BYTES, load_settings

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
    assert s.file_root.endswith("outlook-personal-mcp/files")
    assert s.max_file_bytes == DEFAULT_MAX_FILE_BYTES

def test_allow_permanent_delete_truthy():
    s = load_settings(environ={"OUTLOOK_MCP_CLIENT_ID": "abc",
                               "OUTLOOK_MCP_ALLOW_PERMANENT_DELETE": "true"})
    assert s.allow_permanent_delete is True

def test_file_root_and_limit_from_env(tmp_path):
    root = tmp_path / "files"
    s = load_settings(
        environ={
            "OUTLOOK_MCP_CLIENT_ID": "abc",
            "OUTLOOK_MCP_FILE_ROOT": str(root),
            "OUTLOOK_MCP_MAX_FILE_BYTES": "42",
        }
    )
    assert s.file_root == str(root)
    assert s.max_file_bytes == 42

def test_file_limit_must_be_positive():
    with pytest.raises(ValueError, match="OUTLOOK_MCP_MAX_FILE_BYTES"):
        load_settings(
            environ={
                "OUTLOOK_MCP_CLIENT_ID": "abc",
                "OUTLOOK_MCP_MAX_FILE_BYTES": "0",
            }
        )
