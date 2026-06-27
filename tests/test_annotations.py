"""Tests that MCP tool annotations are correctly applied to every registered tool."""
from __future__ import annotations

import pytest
from outlook_personal_mcp.server import build_server
from outlook_personal_mcp.config import Settings


class StubTokens:
    def get_token(self):
        return "T"


@pytest.fixture
def mcp_no_perm(monkeypatch):
    """Server built without allow_permanent_delete — permanent_delete not registered."""
    from outlook_personal_mcp import server

    monkeypatch.setattr(server, "TokenProvider", lambda *_a, **_k: StubTokens())
    return build_server(Settings(client_id="abc"))


@pytest.fixture
def mcp_with_perm(monkeypatch):
    """Server built with allow_permanent_delete=True — all 32 tools registered."""
    from outlook_personal_mcp import server

    monkeypatch.setattr(server, "TokenProvider", lambda *_a, **_k: StubTokens())
    return build_server(Settings(client_id="abc", allow_permanent_delete=True))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _tool_map(tools):
    return {t.name: t for t in tools}


# ---------------------------------------------------------------------------
# Read-only tools
# ---------------------------------------------------------------------------


async def test_whoami_is_readonly(mcp_no_perm):
    tools = _tool_map(await mcp_no_perm.list_tools())
    ann = tools["whoami"].annotations
    assert ann is not None
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False
    assert ann.openWorldHint is True


async def test_list_messages_is_readonly(mcp_no_perm):
    tools = _tool_map(await mcp_no_perm.list_tools())
    ann = tools["list_messages"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False
    assert ann.openWorldHint is True


async def test_search_messages_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["search_messages"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_get_message_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["get_message"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_list_attachments_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["list_attachments"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_list_folders_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["list_folders"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_list_calendars_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["list_calendars"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_list_events_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["list_events"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_search_events_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["search_events"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_get_event_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["get_event"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


async def test_find_availability_is_readonly(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["find_availability"].annotations
    assert ann.readOnlyHint is True
    assert ann.destructiveHint is False


# ---------------------------------------------------------------------------
# Write non-destructive tools
# ---------------------------------------------------------------------------


async def test_send_mail_is_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["send_mail"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False
    assert ann.openWorldHint is True


async def test_download_attachment_is_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["download_attachment"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False


async def test_create_draft_is_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["create_draft"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False


async def test_update_draft_is_idempotent_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["update_draft"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False
    assert ann.idempotentHint is True


async def test_mark_read_is_idempotent_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["mark_read"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False
    assert ann.idempotentHint is True


async def test_flag_message_is_idempotent_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["flag_message"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False
    assert ann.idempotentHint is True


async def test_move_message_is_idempotent_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["move_message"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False
    assert ann.idempotentHint is True


async def test_create_folder_is_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["create_folder"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False


async def test_create_event_is_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["create_event"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False


async def test_update_event_is_idempotent_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["update_event"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False
    assert ann.idempotentHint is True


async def test_respond_event_is_write(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["respond_event"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is False


# ---------------------------------------------------------------------------
# Destructive tools
# ---------------------------------------------------------------------------


async def test_delete_message_is_destructive(mcp_no_perm):
    tools = _tool_map(await mcp_no_perm.list_tools())
    ann = tools["delete_message"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is True
    assert ann.openWorldHint is True


async def test_delete_folder_is_destructive(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["delete_folder"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is True


async def test_delete_event_is_destructive(mcp_no_perm):
    ann = _tool_map(await mcp_no_perm.list_tools())["delete_event"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is True


async def test_permanent_delete_is_destructive(mcp_with_perm):
    tools = _tool_map(await mcp_with_perm.list_tools())
    assert "permanent_delete" in tools, "permanent_delete not registered with allow_permanent_delete=True"
    ann = tools["permanent_delete"].annotations
    assert ann.readOnlyHint is False
    assert ann.destructiveHint is True
    assert ann.openWorldHint is True


# ---------------------------------------------------------------------------
# Tool count guards
# ---------------------------------------------------------------------------


async def test_tool_count_without_permanent_delete(mcp_no_perm):
    """31 tools without permanent_delete."""
    tools = await mcp_no_perm.list_tools()
    assert len(tools) == 31


async def test_tool_count_with_permanent_delete(mcp_with_perm):
    """32 tools with permanent_delete enabled."""
    tools = await mcp_with_perm.list_tools()
    assert len(tools) == 32


async def test_all_tools_have_annotations(mcp_with_perm):
    """Every registered tool must have a non-None annotations object."""
    tools = await mcp_with_perm.list_tools()
    missing = [t.name for t in tools if t.annotations is None]
    assert missing == [], f"Tools missing annotations: {missing}"


async def test_all_tools_have_title(mcp_with_perm):
    """Every registered tool must have a non-empty title annotation."""
    tools = await mcp_with_perm.list_tools()
    missing = [t.name for t in tools if not (t.annotations and t.annotations.title)]
    assert missing == [], f"Tools missing title annotation: {missing}"


async def test_all_tools_have_open_world_hint(mcp_with_perm):
    """Every tool talks to Microsoft Graph — all must have openWorldHint=True."""
    tools = await mcp_with_perm.list_tools()
    not_open = [t.name for t in tools if not (t.annotations and t.annotations.openWorldHint)]
    assert not_open == [], f"Tools without openWorldHint=True: {not_open}"
