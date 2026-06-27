from __future__ import annotations

import asyncio
import sys
import httpx

BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphError(Exception):
    def __init__(self, status: int, code: str, message: str):
        self.status, self.code, self.message = status, code, message
        super().__init__(f"Graph {status} {code}: {message}")


class GraphClient:
    def __init__(
        self,
        token_provider,
        *,
        base_url: str = BASE_URL,
        max_retries: int = 3,
        backoff_base: float = 0.5,
        timeout: float = 30.0,
        debug: bool = False,
    ):
        self._tokens = token_provider
        self._base = base_url.rstrip("/")
        self._max_retries = max_retries
        self._backoff_base = backoff_base
        self._timeout = timeout
        self._debug = debug

    def _url_for_path(self, path: str) -> str:
        if path.startswith(("http://", "https://")):
            if path == self._base or path.startswith(f"{self._base}/"):
                return path
            raise ValueError("GraphClient absolute URLs must stay under the configured Graph base URL")
        if not path.startswith("/"):
            raise ValueError("GraphClient paths must start with '/'")
        return f"{self._base}{path}"

    async def request(
        self,
        method: str,
        path: str,
        *,
        params=None,
        json=None,
        headers=None,
        raw: bool = False,
    ):
        url = self._url_for_path(path)
        attempt = 0
        async with httpx.AsyncClient(timeout=self._timeout) as http:
            while True:
                token = self._tokens.get_token()
                hdrs = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
                if headers:
                    hdrs.update(headers)
                resp = await http.request(method, url, params=params, json=json, headers=hdrs)
                if self._debug:
                    print(f"[graph] {method} {url} -> {resp.status_code}", file=sys.stderr)
                if resp.status_code in (429, 503) and attempt < self._max_retries:
                    retry_after = float(resp.headers.get("Retry-After", "0") or 0)
                    delay = max(retry_after, self._backoff_base * (2**attempt))
                    await asyncio.sleep(delay)
                    attempt += 1
                    continue
                if resp.status_code >= 400:
                    self._raise(resp)
                if resp.status_code == 204 or not resp.content:
                    return None
                if raw:
                    return resp.content
                return resp.json()

    @staticmethod
    def _raise(resp: httpx.Response) -> None:
        code, message = "unknown", resp.text
        try:
            err = resp.json().get("error", {})
            code = err.get("code", code)
            message = err.get("message", message)
        except Exception:
            pass
        raise GraphError(resp.status_code, code, message)

    async def get(self, path: str, *, params=None, raw: bool = False):
        return await self.request("GET", path, params=params, raw=raw)

    async def post(self, path: str, *, json=None, params=None):
        return await self.request("POST", path, json=json, params=params)

    async def patch(self, path: str, *, json=None):
        return await self.request("PATCH", path, json=json)

    async def delete(self, path: str):
        return await self.request("DELETE", path)
