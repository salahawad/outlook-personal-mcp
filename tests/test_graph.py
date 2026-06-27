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
    respx.get("https://graph.microsoft.com/v1.0/me").mock(
        side_effect=[
            httpx.Response(429, headers={"Retry-After": "0"}, json={"error": {"code": "TooManyRequests"}}),
            httpx.Response(200, json={"id": "1"}),
        ]
    )
    data = await client.get("/me")
    assert data["id"] == "1"


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
