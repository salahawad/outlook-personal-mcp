import httpx
import pytest
import respx
from outlook_personal_mcp.graph import GraphClient, GraphError


class StubTokens:
    def get_token(self):
        return "TESTTOKEN"


@pytest.fixture
def client():
    return GraphClient(StubTokens(), max_retries=2, backoff_base=0)


@respx.mock
async def test_get_injects_bearer_and_parses_json(client):
    route = respx.get("https://graph.microsoft.com/v1.0/me").mock(
        return_value=httpx.Response(200, json={"id": "1", "userPrincipalName": "u@x.com"})
    )
    data = await client.get("/me")
    assert data["userPrincipalName"] == "u@x.com"
    assert route.calls.last.request.headers["Authorization"] == "Bearer TESTTOKEN"


@respx.mock
async def test_retries_on_429_then_succeeds(client):
    route = respx.get("https://graph.microsoft.com/v1.0/me").mock(
        side_effect=[
            httpx.Response(429, headers={"Retry-After": "0"}, json={"error": {"code": "TooManyRequests"}}),
            httpx.Response(200, json={"id": "1"}),
        ]
    )
    data = await client.get("/me")
    assert data["id"] == "1"
    assert route.call_count == 2  # 429 then 200 — proves the retry actually happened


@respx.mock
async def test_maps_error_to_graph_error(client):
    respx.get("https://graph.microsoft.com/v1.0/me").mock(
        return_value=httpx.Response(403, json={"error": {"code": "ErrorAccessDenied", "message": "no scope"}})
    )
    with pytest.raises(GraphError) as ei:
        await client.get("/me")
    assert ei.value.status == 403
    assert ei.value.code == "ErrorAccessDenied"


@respx.mock
async def test_delete_returns_none_on_204(client):
    respx.delete("https://graph.microsoft.com/v1.0/me/messages/1").mock(
        return_value=httpx.Response(204)
    )
    assert await client.delete("/me/messages/1") is None

async def test_rejects_absolute_url_outside_graph_base(client):
    with pytest.raises(ValueError, match="Graph base URL"):
        await client.get("https://example.com/me")

async def test_rejects_relative_path_without_leading_slash(client):
    with pytest.raises(ValueError, match="must start with '/'"):
        await client.get("me")


@respx.mock
async def test_debug_logs_method_url_status_not_token(capsys):
    respx.get("https://graph.microsoft.com/v1.0/me").mock(
        return_value=httpx.Response(200, json={"id": "1"}))
    dbg = GraphClient(StubTokens(), backoff_base=0, debug=True)
    await dbg.get("/me")
    err = capsys.readouterr().err
    assert "GET" in err and "/me" in err and "200" in err
    assert "TESTTOKEN" not in err  # token must never be logged
