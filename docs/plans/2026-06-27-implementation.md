# outlook-personal-mcp Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a generic, public, stdio MCP server in Python that gives Claude Code / Codex full control of a personal Outlook.com mailbox **and** calendar via Microsoft Graph.

**Architecture:** A thin, layered server — `config` (env) → `auth` (MSAL device-code + persistent token cache) → `graph` (httpx REST client with throttling/retry + error mapping) → `tools/*` (FastMCP tool handlers grouped by domain) → `server` (FastMCP wiring, stdio entrypoint). Tool layers call `graph` only; `graph` knows nothing about MCP. Each install is self-owned: the user registers their own Azure app and signs in as themselves.

**Tech Stack:** Python 3.10+, `mcp` (FastMCP), `msal`, `httpx`, `pydantic`; dev: `pytest`, `respx`, `ruff`; packaging with `uv`/hatchling; GitHub Actions CI.

**Reference:** [docs/DESIGN.md](../DESIGN.md). Microsoft Graph v1.0 docs and MSAL Python docs should be consulted by the implementer for any endpoint flagged "verify".

---

## File Structure

| File | Responsibility |
| --- | --- |
| `pyproject.toml` | Build config, deps, console script `outlook-personal-mcp`, ruff/pytest config. |
| `src/outlook_personal_mcp/__init__.py` | Package version. |
| `src/outlook_personal_mcp/__main__.py` | `python -m outlook_personal_mcp` → CLI dispatch (`serve` default, `login`). |
| `src/outlook_personal_mcp/config.py` | `Settings` frozen dataclass loaded from env. |
| `src/outlook_personal_mcp/auth.py` | `TokenProvider`: MSAL device-code login, file token cache, silent refresh. |
| `src/outlook_personal_mcp/graph.py` | `GraphClient`: async httpx wrapper, retry/throttle, error mapping, typed `GraphError`. |
| `src/outlook_personal_mcp/models.py` | Pydantic input/output models + output-shaping helpers. |
| `src/outlook_personal_mcp/tools/__init__.py` | `register_all(mcp, client)` aggregator. |
| `src/outlook_personal_mcp/tools/mail.py` | Read/search/send/reply/forward/organize/delete message tools. |
| `src/outlook_personal_mcp/tools/folders.py` | Mail folder tools. |
| `src/outlook_personal_mcp/tools/drafts.py` | Draft create/update/send + attachments. |
| `src/outlook_personal_mcp/tools/calendar.py` | Event + availability tools. |
| `src/outlook_personal_mcp/server.py` | Build FastMCP app, register tools, `whoami`, run stdio. |
| `tests/conftest.py` | Fixtures: fake settings, respx mock, in-memory token provider. |
| `tests/test_*.py` | Unit tests (HTTP fully mocked; no live calls). |
| `README.md`, `.env.example`, `LICENSE`, `.github/workflows/ci.yml` | Docs, sample env, license, CI. |

---

## Conventions for every code task (TDD loop)

Each task repeats this loop. The standard commands:

- Run a single test: `uv run pytest tests/<file>::<test> -v`
- Run all tests: `uv run pytest -q`
- Lint: `uv run ruff check .`
- Commit: `git add -A && git commit -m "<msg>"`

Author commits as `salahawad <salah.awad@outlook.com>` (already set as repo git identity).

---

## Task 1: Project scaffolding

**Files:**
- Create: `pyproject.toml`, `src/outlook_personal_mcp/__init__.py`, `.gitignore`, `LICENSE`, `.env.example`, `tests/__init__.py`

- [ ] **Step 1: Create `.gitignore`**

```gitignore
__pycache__/
*.py[cod]
.venv/
dist/
build/
*.egg-info/
.pytest_cache/
.ruff_cache/
.env
token_cache.bin
```

- [ ] **Step 2: Create `LICENSE` (MIT)**

Standard MIT text, `Copyright (c) 2026 Salah Awad`.

- [ ] **Step 3: Create `pyproject.toml`**

```toml
[project]
name = "outlook-personal-mcp"
version = "0.1.0"
description = "MCP server for a personal Outlook.com mailbox and calendar via Microsoft Graph"
readme = "README.md"
requires-python = ">=3.10"
license = { text = "MIT" }
authors = [{ name = "Salah Awad" }]
dependencies = [
  "mcp>=1.2.0",
  "msal>=1.28.0",
  "httpx>=0.27.0",
  "pydantic>=2.6.0",
]

[project.scripts]
outlook-personal-mcp = "outlook_personal_mcp.__main__:main"

[project.optional-dependencies]
dev = ["pytest>=8.0", "pytest-asyncio>=0.23", "respx>=0.21", "ruff>=0.4"]

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/outlook_personal_mcp"]

[tool.pytest.ini_options]
asyncio_mode = "auto"
testpaths = ["tests"]

[tool.ruff]
line-length = 100
target-version = "py310"
```

- [ ] **Step 4: Create `src/outlook_personal_mcp/__init__.py`**

```python
__version__ = "0.1.0"
```

- [ ] **Step 5: Create `tests/__init__.py`** (empty file) and `.env.example`

```bash
# .env.example
OUTLOOK_MCP_CLIENT_ID=your-azure-app-client-id
# OUTLOOK_MCP_AUTHORITY=https://login.microsoftonline.com/consumers
# OUTLOOK_MCP_TOKEN_CACHE=~/.config/outlook-personal-mcp/token_cache.bin
# OUTLOOK_MCP_ALLOW_PERMANENT_DELETE=false
# OUTLOOK_MCP_DEBUG=false
```

- [ ] **Step 6: Create the venv and install**

Run: `uv venv && uv pip install -e ".[dev]"`
Expected: installs without error.

- [ ] **Step 7: Commit**

```bash
git add -A && git commit -m "chore: project scaffolding (pyproject, license, gitignore)"
```

---

## Task 2: `config.py` — settings from env

**Files:**
- Create: `src/outlook_personal_mcp/config.py`, `tests/test_config.py`

- [ ] **Step 1: Write the failing test** — `tests/test_config.py`

```python
import os
import pytest
from outlook_personal_mcp.config import Settings, load_settings

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
```

- [ ] **Step 2: Run to verify it fails**

Run: `uv run pytest tests/test_config.py -v`
Expected: FAIL (module/attr missing).

- [ ] **Step 3: Implement `config.py`**

```python
from __future__ import annotations
import os
from dataclasses import dataclass, field

DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"
# offline_access is added by MSAL automatically; do not list it here.
SCOPES = ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "User.Read"]

def _default_cache_path() -> str:
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, "outlook-personal-mcp", "token_cache.bin")

def _truthy(v: str | None) -> bool:
    return str(v).strip().lower() in {"1", "true", "yes", "on"}

@dataclass(frozen=True)
class Settings:
    client_id: str
    authority: str = DEFAULT_AUTHORITY
    scopes: list[str] = field(default_factory=lambda: list(SCOPES))
    token_cache_path: str = field(default_factory=_default_cache_path)
    allow_permanent_delete: bool = False
    debug: bool = False

def load_settings(environ: dict | None = None) -> Settings:
    env = os.environ if environ is None else environ
    client_id = env.get("OUTLOOK_MCP_CLIENT_ID", "").strip()
    if not client_id:
        raise ValueError("OUTLOOK_MCP_CLIENT_ID is required")
    cache = env.get("OUTLOOK_MCP_TOKEN_CACHE")
    return Settings(
        client_id=client_id,
        authority=env.get("OUTLOOK_MCP_AUTHORITY", DEFAULT_AUTHORITY),
        token_cache_path=os.path.expanduser(cache) if cache else _default_cache_path(),
        allow_permanent_delete=_truthy(env.get("OUTLOOK_MCP_ALLOW_PERMANENT_DELETE")),
        debug=_truthy(env.get("OUTLOOK_MCP_DEBUG")),
    )
```

- [ ] **Step 4: Run to verify it passes**

Run: `uv run pytest tests/test_config.py -v` → Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(config): env-driven Settings with sane defaults"
```

---

## Task 3: `auth.py` — MSAL device-code + token cache

**Files:**
- Create: `src/outlook_personal_mcp/auth.py`, `tests/test_auth.py`

**Notes:** MSAL `PublicClientApplication` with a `SerializableTokenCache` persisted to disk (chmod 600). `get_token()` tries `acquire_token_silent` first; if no account/refresh works, raises `AuthRequired` carrying the device-flow message so the caller can surface it. `login()` runs the blocking device flow (used by the `login` CLI and lazily on first call when interactive).

- [ ] **Step 1: Write the failing test** — `tests/test_auth.py`

```python
import json
import pytest
from outlook_personal_mcp.auth import TokenProvider, AuthRequired
from outlook_personal_mcp.config import Settings

class FakeMsalApp:
    def __init__(self, *, silent=None, flow=None, by_flow=None, accounts=None):
        self._silent, self._flow, self._by_flow = silent, flow, by_flow
        self._accounts = accounts or []
    def get_accounts(self): return self._accounts
    def acquire_token_silent(self, scopes, account): return self._silent
    def initiate_device_flow(self, scopes): return self._flow
    def acquire_token_by_device_flow(self, flow): return self._by_flow

def _settings(tmp_path):
    return Settings(client_id="abc", token_cache_path=str(tmp_path / "cache.bin"))

def test_silent_token_returns_access_token(tmp_path):
    app = FakeMsalApp(silent={"access_token": "TOK"}, accounts=[{"username": "u"}])
    tp = TokenProvider(_settings(tmp_path), app_factory=lambda *_a, **_k: app)
    assert tp.get_token() == "TOK"

def test_no_account_raises_auth_required(tmp_path):
    app = FakeMsalApp(silent=None, accounts=[],
                      flow={"verification_uri": "https://aka.ms/devicelogin",
                            "user_code": "ABCD-1234", "message": "Go authenticate"})
    tp = TokenProvider(_settings(tmp_path), app_factory=lambda *_a, **_k: app)
    with pytest.raises(AuthRequired) as ei:
        tp.get_token()
    assert "ABCD-1234" in str(ei.value)

def test_login_persists_cache(tmp_path):
    app = FakeMsalApp(flow={"user_code": "X", "message": "m"},
                      by_flow={"access_token": "TOK", "token_type": "Bearer"})
    s = _settings(tmp_path)
    tp = TokenProvider(s, app_factory=lambda *_a, **_k: app)
    tp.login()  # should not raise
    # cache file created with restrictive perms
    import os, stat
    assert os.path.exists(s.token_cache_path)
    assert stat.S_IMODE(os.stat(s.token_cache_path).st_mode) == 0o600
```

- [ ] **Step 2: Run to verify it fails**

Run: `uv run pytest tests/test_auth.py -v` → Expected: FAIL.

- [ ] **Step 3: Implement `auth.py`**

```python
from __future__ import annotations
import os
import sys
import msal
from .config import Settings

class AuthRequired(Exception):
    """Raised when interactive device-code login is needed. str() is user-facing."""

class TokenProvider:
    def __init__(self, settings: Settings, app_factory=None):
        self._s = settings
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        factory = app_factory or self._default_app_factory
        self._app = factory(settings.client_id, authority=settings.authority,
                            token_cache=self._cache)

    @staticmethod
    def _default_app_factory(client_id, *, authority, token_cache):
        return msal.PublicClientApplication(client_id, authority=authority,
                                            token_cache=token_cache)

    def _load_cache(self):
        path = self._s.token_cache_path
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as fh:
                self._cache.deserialize(fh.read())

    def _save_cache(self):
        if not self._cache.has_state_changed:
            return
        path = self._s.token_cache_path
        os.makedirs(os.path.dirname(path), exist_ok=True)
        # Write then tighten perms to 600.
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(self._cache.serialize())
        os.chmod(path, 0o600)

    def get_token(self) -> str:
        accounts = self._app.get_accounts()
        if accounts:
            res = self._app.acquire_token_silent(self._s.scopes, account=accounts[0])
            if res and "access_token" in res:
                self._save_cache()
                return res["access_token"]
        flow = self._app.initiate_device_flow(scopes=self._s.scopes)
        msg = flow.get("message", "Sign in to continue.")
        raise AuthRequired(
            f"Authentication required. {msg}\n"
            "Run `outlook-personal-mcp login` (or retry this tool after signing in)."
        )

    def login(self) -> None:
        flow = self._app.initiate_device_flow(scopes=self._s.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device flow: {flow.get('error_description', flow)}")
        print(flow["message"], file=sys.stderr, flush=True)
        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Login failed: {result.get('error_description', result)}")
        self._save_cache()
```

- [ ] **Step 4: Run to verify it passes**

Run: `uv run pytest tests/test_auth.py -v` → Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(auth): MSAL device-code login with 600-perm token cache"
```

---

## Task 4: `graph.py` — httpx Graph client with retry + error mapping

**Files:**
- Create: `src/outlook_personal_mcp/graph.py`, `tests/test_graph.py`

**Notes:** Async client. Injects bearer from `TokenProvider.get_token()`. Honors `Retry-After` on 429/503 with capped retries. Maps non-2xx to `GraphError(status, code, message)`. Base URL `https://graph.microsoft.com/v1.0`. Tokens never logged.

- [ ] **Step 1: Write the failing test** — `tests/test_graph.py`

```python
import httpx, pytest, respx
from outlook_personal_mcp.graph import GraphClient, GraphError

class StubTokens:
    def get_token(self): return "TESTTOKEN"

@pytest.fixture
def client():
    return GraphClient(StubTokens(), max_retries=2, backoff_base=0)

@respx.mock
async def test_get_injects_bearer_and_parses_json(client):
    route = respx.get("https://graph.microsoft.com/v1.0/me").mock(
        return_value=httpx.Response(200, json={"id": "1", "userPrincipalName": "u@x.com"}))
    data = await client.get("/me")
    assert data["userPrincipalName"] == "u@x.com"
    assert route.calls.last.request.headers["Authorization"] == "Bearer TESTTOKEN"

@respx.mock
async def test_retries_on_429_then_succeeds(client):
    respx.get("https://graph.microsoft.com/v1.0/me").mock(side_effect=[
        httpx.Response(429, headers={"Retry-After": "0"}, json={"error": {"code": "TooManyRequests"}}),
        httpx.Response(200, json={"id": "1"}),
    ])
    data = await client.get("/me")
    assert data["id"] == "1"

@respx.mock
async def test_maps_error_to_graph_error(client):
    respx.get("https://graph.microsoft.com/v1.0/me").mock(
        return_value=httpx.Response(403, json={"error": {"code": "ErrorAccessDenied",
                                                          "message": "no scope"}}))
    with pytest.raises(GraphError) as ei:
        await client.get("/me")
    assert ei.value.status == 403
    assert ei.value.code == "ErrorAccessDenied"

@respx.mock
async def test_delete_returns_none_on_204(client):
    respx.delete("https://graph.microsoft.com/v1.0/me/messages/1").mock(
        return_value=httpx.Response(204))
    assert await client.delete("/me/messages/1") is None
```

- [ ] **Step 2: Run to verify it fails**

Run: `uv run pytest tests/test_graph.py -v` → Expected: FAIL.

- [ ] **Step 3: Implement `graph.py`**

```python
from __future__ import annotations
import asyncio
import httpx

BASE_URL = "https://graph.microsoft.com/v1.0"

class GraphError(Exception):
    def __init__(self, status: int, code: str, message: str):
        self.status, self.code, self.message = status, code, message
        super().__init__(f"Graph {status} {code}: {message}")

class GraphClient:
    def __init__(self, token_provider, *, base_url: str = BASE_URL,
                 max_retries: int = 3, backoff_base: float = 0.5, timeout: float = 30.0):
        self._tokens = token_provider
        self._base = base_url.rstrip("/")
        self._max_retries = max_retries
        self._backoff_base = backoff_base
        self._timeout = timeout

    async def request(self, method: str, path: str, *, params=None, json=None,
                      headers=None, raw=False):
        url = path if path.startswith("http") else f"{self._base}{path}"
        attempt = 0
        async with httpx.AsyncClient(timeout=self._timeout) as http:
            while True:
                token = self._tokens.get_token()
                hdrs = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
                if headers:
                    hdrs.update(headers)
                resp = await http.request(method, url, params=params, json=json, headers=hdrs)
                if resp.status_code in (429, 503) and attempt < self._max_retries:
                    retry_after = float(resp.headers.get("Retry-After", "0") or 0)
                    await asyncio.sleep(max(retry_after, self._backoff_base * (2 ** attempt)))
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
    def _raise(resp: httpx.Response):
        code, message = "unknown", resp.text
        try:
            err = resp.json().get("error", {})
            code = err.get("code", code)
            message = err.get("message", message)
        except Exception:
            pass
        raise GraphError(resp.status_code, code, message)

    async def get(self, path, *, params=None, raw=False):
        return await self.request("GET", path, params=params, raw=raw)

    async def post(self, path, *, json=None, params=None):
        return await self.request("POST", path, json=json, params=params)

    async def patch(self, path, *, json=None):
        return await self.request("PATCH", path, json=json)

    async def delete(self, path):
        return await self.request("DELETE", path)
```

- [ ] **Step 4: Run to verify it passes**

Run: `uv run pytest tests/test_graph.py -v` → Expected: PASS (all 4).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(graph): async httpx Graph client with throttle-retry + error mapping"
```

---

## Task 5: `models.py` — pydantic schemas + output shapers

**Files:**
- Create: `src/outlook_personal_mcp/models.py`, `tests/test_models.py`

**Notes:** Input models give clean MCP tool schemas; output shapers trim Graph payloads to compact dicts to save tokens.

- [ ] **Step 1: Write the failing test** — `tests/test_models.py`

```python
from outlook_personal_mcp.models import (Recipient, SendMailInput, shape_message, shape_event)

def test_send_mail_input_normalizes_recipients():
    m = SendMailInput(to=["a@x.com"], subject="hi", body="hello")
    payload = m.to_graph()
    assert payload["message"]["toRecipients"] == [{"emailAddress": {"address": "a@x.com"}}]
    assert payload["message"]["body"]["contentType"] == "Text"
    assert payload["message"]["subject"] == "hi"

def test_shape_message_trims_fields():
    raw = {"id": "1", "subject": "S", "isRead": False,
           "from": {"emailAddress": {"address": "a@x.com", "name": "A"}},
           "receivedDateTime": "2026-06-27T10:00:00Z", "bodyPreview": "p",
           "hasAttachments": True, "extra": "drop me"}
    s = shape_message(raw)
    assert s == {"id": "1", "subject": "S", "from": "a@x.com", "is_read": False,
                 "received": "2026-06-27T10:00:00Z", "preview": "p", "has_attachments": True}

def test_shape_event_trims_fields():
    raw = {"id": "e1", "subject": "Sync", "start": {"dateTime": "2026-06-27T09:00:00", "timeZone": "UTC"},
           "end": {"dateTime": "2026-06-27T09:30:00", "timeZone": "UTC"},
           "location": {"displayName": "Room"}, "isAllDay": False, "junk": 1}
    s = shape_event(raw)
    assert s["id"] == "e1" and s["subject"] == "Sync"
    assert s["start"] == "2026-06-27T09:00:00 UTC" and s["location"] == "Room"
```

- [ ] **Step 2: Run to verify it fails** → `uv run pytest tests/test_models.py -v` → FAIL.

- [ ] **Step 3: Implement `models.py`**

```python
from __future__ import annotations
from typing import Literal, Optional
from pydantic import BaseModel, Field

Recipient = str  # plain email address

def _recipients(addrs: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in addrs]

class SendMailInput(BaseModel):
    to: list[str]
    subject: str
    body: str
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    html: bool = False
    save_to_sent: bool = True

    def to_graph(self) -> dict:
        msg = {
            "subject": self.subject,
            "body": {"contentType": "HTML" if self.html else "Text", "content": self.body},
            "toRecipients": _recipients(self.to),
        }
        if self.cc:
            msg["ccRecipients"] = _recipients(self.cc)
        if self.bcc:
            msg["bccRecipients"] = _recipients(self.bcc)
        return {"message": msg, "saveToSentItems": self.save_to_sent}

class EventInput(BaseModel):
    subject: str
    start: str  # ISO 8601, e.g. 2026-06-27T09:00:00
    end: str
    time_zone: str = "UTC"
    attendees: list[str] = Field(default_factory=list)
    location: Optional[str] = None
    body: Optional[str] = None
    is_online_meeting: bool = False

    def to_graph(self) -> dict:
        ev: dict = {
            "subject": self.subject,
            "start": {"dateTime": self.start, "timeZone": self.time_zone},
            "end": {"dateTime": self.end, "timeZone": self.time_zone},
        }
        if self.attendees:
            ev["attendees"] = [{"emailAddress": {"address": a}, "type": "required"}
                               for a in self.attendees]
        if self.location:
            ev["location"] = {"displayName": self.location}
        if self.body:
            ev["body"] = {"contentType": "HTML", "content": self.body}
        if self.is_online_meeting:
            ev["isOnlineMeeting"] = True
            ev["onlineMeetingProvider"] = "teamsForBusiness"
        return ev

def shape_message(m: dict) -> dict:
    return {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "is_read": m.get("isRead"),
        "received": m.get("receivedDateTime"),
        "preview": m.get("bodyPreview"),
        "has_attachments": m.get("hasAttachments"),
    }

def shape_event(e: dict) -> dict:
    def _dt(d):
        d = d or {}
        return f"{d.get('dateTime')} {d.get('timeZone')}".strip()
    return {
        "id": e.get("id"),
        "subject": e.get("subject"),
        "start": _dt(e.get("start")),
        "end": _dt(e.get("end")),
        "location": (e.get("location") or {}).get("displayName"),
        "is_all_day": e.get("isAllDay"),
        "organizer": (e.get("organizer") or {}).get("emailAddress", {}).get("address"),
    }
```

- [ ] **Step 4: Run to verify it passes** → `uv run pytest tests/test_models.py -v` → PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(models): pydantic inputs + compact output shapers"
```

---

## Task 6: `server.py` + `__main__.py` — FastMCP wiring, `whoami`, CLI

**Files:**
- Create: `src/outlook_personal_mcp/server.py`, `src/outlook_personal_mcp/tools/__init__.py`, `src/outlook_personal_mcp/__main__.py`, `tests/test_server.py`

**Notes:** `build_server(settings)` constructs `TokenProvider`, `GraphClient`, a FastMCP instance, registers `whoami`, then calls `tools.register_all`. The CLI: `serve` (default) runs stdio; `login` runs the device flow.

- [ ] **Step 1: Write the failing test** — `tests/test_server.py`

```python
import respx, httpx, pytest
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

async def test_whoami_registered_and_calls_me(mcp):
    tools = await mcp.list_tools()
    names = {t.name for t in tools}
    assert "whoami" in names
    assert "send_mail" in names      # proves register_all wired
    assert "list_events" in names    # proves calendar wired
```

- [ ] **Step 2: Run to verify it fails** → `uv run pytest tests/test_server.py -v` → FAIL.

- [ ] **Step 3: Implement `tools/__init__.py`** (stub aggregator that grows per task)

```python
from . import mail, folders, drafts, calendar

def register_all(mcp, client, settings):
    folders.register(mcp, client)
    mail.register(mcp, client, settings)
    drafts.register(mcp, client)
    calendar.register(mcp, client)
```

- [ ] **Step 4: Implement `server.py`**

```python
from __future__ import annotations
from mcp.server.fastmcp import FastMCP
from .auth import TokenProvider
from .graph import GraphClient
from .config import Settings, load_settings
from . import tools

def build_server(settings: Settings) -> FastMCP:
    tokens = TokenProvider(settings)
    client = GraphClient(tokens)
    mcp = FastMCP("outlook-personal-mcp")

    @mcp.tool()
    async def whoami() -> dict:
        """Return the signed-in user's Microsoft account profile."""
        me = await client.get("/me")
        return {"display_name": me.get("displayName"),
                "email": me.get("userPrincipalName") or me.get("mail"),
                "id": me.get("id")}

    tools.register_all(mcp, client, settings)
    return mcp

def run():
    settings = load_settings()
    build_server(settings).run()  # stdio transport
```

- [ ] **Step 5: Implement `__main__.py`**

```python
from __future__ import annotations
import sys
from .config import load_settings
from .auth import TokenProvider
from .server import build_server

def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    cmd = argv[0] if argv else "serve"
    settings = load_settings()
    if cmd == "login":
        TokenProvider(settings).login()
        print("Login successful; token cached.", file=sys.stderr)
        return 0
    if cmd in ("serve", ""):
        build_server(settings).run()
        return 0
    print(f"Unknown command: {cmd}. Use 'serve' or 'login'.", file=sys.stderr)
    return 2

if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 6:** Create minimal empty `tools/mail.py`, `tools/folders.py`, `tools/drafts.py`, `tools/calendar.py`, each with `def register(*_a, **_k): pass` so imports resolve. (Tasks 7–13 replace these.)

- [ ] **Step 7: Run to verify it passes** → `uv run pytest tests/test_server.py -v`
  Expected: FAIL on `send_mail`/`list_events` asserts (registers are stubs). Reduce the assert in this task to only `"whoami" in names`, get PASS, then commit. Re-add the broader asserts in Task 13's final step.

- [ ] **Step 8: Commit**

```bash
git add -A && git commit -m "feat(server): FastMCP wiring, whoami tool, serve/login CLI"
```

---

## Task 7: `tools/folders.py` — mail folder tools

**Files:**
- Create (replace stub): `src/outlook_personal_mcp/tools/folders.py`, `tests/test_tools_folders.py`

**Graph endpoints:** list `GET /me/mailFolders?$top=100`; create `POST /me/mailFolders {displayName}`; rename `PATCH /me/mailFolders/{id} {displayName}`; delete `DELETE /me/mailFolders/{id}`.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_folders.py`

```python
import httpx, respx, pytest
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import folders

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    client = GraphClient(StubTokens(), backoff_base=0)
    folders.register(mcp, client)
    return mcp

async def _call(mcp, name, **args):
    return await mcp.call_tool(name, args)

@respx.mock
async def test_list_folders(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/mailFolders").mock(
        return_value=httpx.Response(200, json={"value": [
            {"id": "f1", "displayName": "Inbox", "unreadItemCount": 3, "totalItemCount": 10}]}))
    result = await _call(ctx, "list_folders")
    # FastMCP returns (content, structured) — assert on structured payload
    assert any("Inbox" in str(c) for c in result[0])

@respx.mock
async def test_create_folder(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/mailFolders").mock(
        return_value=httpx.Response(201, json={"id": "f2", "displayName": "Projects"}))
    await _call(ctx, "create_folder", name="Projects")
    assert route.called
    assert route.calls.last.request.read() == b'{"displayName": "Projects"}'
```

- [ ] **Step 2: Run to verify it fails** → `uv run pytest tests/test_tools_folders.py -v` → FAIL.

- [ ] **Step 3: Implement `tools/folders.py`**

```python
from __future__ import annotations

def register(mcp, client):
    @mcp.tool()
    async def list_folders() -> list[dict]:
        """List mail folders with unread/total counts."""
        data = await client.get("/me/mailFolders", params={"$top": 100})
        return [{"id": f["id"], "name": f["displayName"],
                 "unread": f.get("unreadItemCount"), "total": f.get("totalItemCount")}
                for f in data.get("value", [])]

    @mcp.tool()
    async def create_folder(name: str, parent_folder_id: str | None = None) -> dict:
        """Create a mail folder, optionally nested under parent_folder_id."""
        path = (f"/me/mailFolders/{parent_folder_id}/childFolders"
                if parent_folder_id else "/me/mailFolders")
        f = await client.post(path, json={"displayName": name})
        return {"id": f["id"], "name": f["displayName"]}

    @mcp.tool()
    async def rename_folder(folder_id: str, new_name: str) -> dict:
        """Rename a mail folder."""
        f = await client.patch(f"/me/mailFolders/{folder_id}", json={"displayName": new_name})
        return {"id": f["id"], "name": f["displayName"]}

    @mcp.tool()
    async def delete_folder(folder_id: str) -> dict:
        """Delete a mail folder (moves it to Deleted Items)."""
        await client.delete(f"/me/mailFolders/{folder_id}")
        return {"deleted": folder_id}
```

- [ ] **Step 4: Run to verify it passes** → `uv run pytest tests/test_tools_folders.py -v` → PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(tools): mail folder tools (list/create/rename/delete)"
```

---

## Task 8: `tools/mail.py` (part 1) — read/search/get/attachments

**Files:**
- Create (replace stub): `src/outlook_personal_mcp/tools/mail.py`, `tests/test_tools_mail_read.py`

**Graph endpoints:** list `GET /me/messages` or `GET /me/mailFolders/{id}/messages` with `$top`,`$skip`,`$filter`,`$select`,`$orderby`; search `GET /me/messages?$search="kql"`; get `GET /me/messages/{id}` (`$select` body when full); attachments list `GET /me/messages/{id}/attachments`; download `GET /me/messages/{id}/attachments/{aid}/$value` (raw bytes).

**Note:** `register(mcp, client, settings)` — `settings` needed later for the permanent-delete gate (Task 11). Define the function signature now.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_mail_read.py`

```python
import httpx, respx, pytest
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    client = GraphClient(StubTokens(), backoff_base=0)
    mail.register(mcp, client, Settings(client_id="abc"))
    return mcp

@respx.mock
async def test_list_messages_shapes_output(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(200, json={"value": [
            {"id": "m1", "subject": "Hi", "isRead": False,
             "from": {"emailAddress": {"address": "a@x.com"}},
             "receivedDateTime": "2026-06-27T10:00:00Z", "bodyPreview": "p",
             "hasAttachments": False}]}))
    res = await ctx.call_tool("list_messages", {"top": 5})
    assert any("Hi" in str(c) for c in res[0])

@respx.mock
async def test_search_messages_passes_search_param(ctx):
    route = respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(200, json={"value": []}))
    await ctx.call_tool("search_messages", {"query": "invoice"})
    assert route.calls.last.request.url.params["$search"] == '"invoice"'
```

- [ ] **Step 2: Run to verify it fails** → FAIL.

- [ ] **Step 3: Implement `tools/mail.py` (part 1 — read section; later tasks append in the same `register`)**

```python
from __future__ import annotations
import base64
from ..models import shape_message
from ..graph import GraphError

_LIST_SELECT = "id,subject,from,isRead,receivedDateTime,bodyPreview,hasAttachments"

def register(mcp, client, settings):
    @mcp.tool()
    async def list_messages(folder_id: str | None = None, top: int = 25,
                            skip: int = 0, unread_only: bool = False) -> list[dict]:
        """List messages (newest first). Optionally restrict to a folder or unread only."""
        path = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
        params = {"$top": top, "$skip": skip, "$select": _LIST_SELECT,
                  "$orderby": "receivedDateTime desc"}
        if unread_only:
            params["$filter"] = "isRead eq false"
        data = await client.get(path, params=params)
        return [shape_message(m) for m in data.get("value", [])]

    @mcp.tool()
    async def search_messages(query: str, top: int = 25) -> list[dict]:
        """Full-text search across the mailbox (Graph $search)."""
        params = {"$search": f'"{query}"', "$top": top, "$select": _LIST_SELECT}
        data = await client.get("/me/messages", params=params)
        return [shape_message(m) for m in data.get("value", [])]

    @mcp.tool()
    async def get_message(message_id: str, include_body: bool = True) -> dict:
        """Get a single message; include_body returns the full body text/HTML."""
        sel = _LIST_SELECT + (",body,toRecipients,ccRecipients" if include_body else "")
        m = await client.get(f"/me/messages/{message_id}", params={"$select": sel})
        out = shape_message(m)
        if include_body:
            out["body"] = (m.get("body") or {}).get("content")
            out["to"] = [r["emailAddress"]["address"] for r in m.get("toRecipients", [])]
        return out

    @mcp.tool()
    async def list_attachments(message_id: str) -> list[dict]:
        """List a message's attachments (id, name, size, contentType)."""
        data = await client.get(f"/me/messages/{message_id}/attachments",
                                params={"$select": "id,name,size,contentType"})
        return [{"id": a["id"], "name": a.get("name"), "size": a.get("size"),
                 "content_type": a.get("contentType")} for a in data.get("value", [])]

    @mcp.tool()
    async def download_attachment(message_id: str, attachment_id: str, save_path: str) -> dict:
        """Download an attachment to a local file path."""
        raw = await client.get(
            f"/me/messages/{message_id}/attachments/{attachment_id}/$value", raw=True)
        with open(save_path, "wb") as fh:
            fh.write(raw)
        return {"saved": save_path, "bytes": len(raw)}

    # Tasks 9 and 11 append more @mcp.tool() handlers inside this same register().
    _register_write(mcp, client)
    _register_organize(mcp, client, settings)
```

> Implementation note: declare `_register_write` and `_register_organize` as module-level functions added in Tasks 9 and 11. For Task 8's green test, temporarily define them as `def _register_write(*a): pass` / `def _register_organize(*a): pass` at module bottom, then flesh out in those tasks.

- [ ] **Step 4: Run to verify it passes** → PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(tools): mail read/search/get/attachment tools"
```

---

## Task 9: `tools/mail.py` (part 2) — send/reply/forward

**Files:**
- Modify: `src/outlook_personal_mcp/tools/mail.py` (fill `_register_write`)
- Test: `tests/test_tools_mail_send.py`

**Graph endpoints:** send `POST /me/sendMail`; reply `POST /me/messages/{id}/reply {comment}`; reply-all `POST /me/messages/{id}/replyAll {comment}`; forward `POST /me/messages/{id}/forward {comment, toRecipients}`.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_mail_send.py`

```python
import httpx, respx, pytest, json
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0), Settings(client_id="abc"))
    return mcp

@respx.mock
async def test_send_mail_builds_graph_payload(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/sendMail").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("send_mail", {"to": ["a@x.com"], "subject": "S", "body": "B"})
    body = json.loads(route.calls.last.request.read())
    assert body["message"]["toRecipients"][0]["emailAddress"]["address"] == "a@x.com"

@respx.mock
async def test_forward_passes_recipients(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/m1/forward").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("forward", {"message_id": "m1", "to": ["b@x.com"], "comment": "fyi"})
    body = json.loads(route.calls.last.request.read())
    assert body["toRecipients"][0]["emailAddress"]["address"] == "b@x.com"
    assert body["comment"] == "fyi"
```

- [ ] **Step 2: Run to verify it fails** → FAIL.

- [ ] **Step 3: Implement `_register_write` in `tools/mail.py`**

```python
from ..models import SendMailInput

def _register_write(mcp, client):
    @mcp.tool()
    async def send_mail(to: list[str], subject: str, body: str,
                        cc: list[str] | None = None, bcc: list[str] | None = None,
                        html: bool = False) -> dict:
        """Send an email."""
        payload = SendMailInput(to=to, subject=subject, body=body,
                                cc=cc or [], bcc=bcc or [], html=html).to_graph()
        await client.post("/me/sendMail", json=payload)
        return {"sent": True, "to": to, "subject": subject}

    @mcp.tool()
    async def reply(message_id: str, comment: str, reply_all: bool = False) -> dict:
        """Reply to a message (set reply_all to reply to everyone)."""
        action = "replyAll" if reply_all else "reply"
        await client.post(f"/me/messages/{message_id}/{action}", json={"comment": comment})
        return {"replied": message_id, "reply_all": reply_all}

    @mcp.tool()
    async def forward(message_id: str, to: list[str], comment: str = "") -> dict:
        """Forward a message to recipients with an optional comment."""
        payload = {"comment": comment,
                   "toRecipients": [{"emailAddress": {"address": a}} for a in to]}
        await client.post(f"/me/messages/{message_id}/forward", json=payload)
        return {"forwarded": message_id, "to": to}
```

- [ ] **Step 4: Run to verify it passes** → PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(tools): mail send/reply/replyAll/forward tools"
```

---

## Task 10: `tools/drafts.py` — drafts + attachments

**Files:**
- Create (replace stub): `src/outlook_personal_mcp/tools/drafts.py`, `tests/test_tools_drafts.py`

**Graph endpoints:** create draft `POST /me/messages {message-resource}`; update `PATCH /me/messages/{id}`; send draft `POST /me/messages/{id}/send`; add attachment `POST /me/messages/{id}/attachments` with `{"@odata.type": "#microsoft.graph.fileAttachment", name, contentBytes(base64)}`.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_drafts.py`

```python
import httpx, respx, pytest, json
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import drafts

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    drafts.register(mcp, GraphClient(StubTokens(), backoff_base=0))
    return mcp

@respx.mock
async def test_create_draft(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages").mock(
        return_value=httpx.Response(201, json={"id": "d1"}))
    res = await ctx.call_tool("create_draft", {"to": ["a@x.com"], "subject": "S", "body": "B"})
    body = json.loads(route.calls.last.request.read())
    assert body["subject"] == "S"
    assert any("d1" in str(c) for c in res[0])

@respx.mock
async def test_send_draft(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/d1/send").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("send_draft", {"draft_id": "d1"})
    assert route.called
```

- [ ] **Step 2: Run to verify it fails** → FAIL.

- [ ] **Step 3: Implement `tools/drafts.py`**

```python
from __future__ import annotations
import base64, os

def _recips(addrs):
    return [{"emailAddress": {"address": a}} for a in addrs]

def register(mcp, client):
    @mcp.tool()
    async def create_draft(to: list[str], subject: str, body: str,
                           cc: list[str] | None = None, html: bool = False) -> dict:
        """Create a draft message (not sent)."""
        msg = {"subject": subject,
               "body": {"contentType": "HTML" if html else "Text", "content": body},
               "toRecipients": _recips(to)}
        if cc:
            msg["ccRecipients"] = _recips(cc)
        d = await client.post("/me/messages", json=msg)
        return {"id": d["id"], "subject": subject}

    @mcp.tool()
    async def update_draft(draft_id: str, subject: str | None = None,
                           body: str | None = None, html: bool = False) -> dict:
        """Update a draft's subject and/or body."""
        patch: dict = {}
        if subject is not None:
            patch["subject"] = subject
        if body is not None:
            patch["body"] = {"contentType": "HTML" if html else "Text", "content": body}
        d = await client.patch(f"/me/messages/{draft_id}", json=patch)
        return {"id": d["id"]}

    @mcp.tool()
    async def add_attachment(draft_id: str, file_path: str) -> dict:
        """Attach a local file to a draft."""
        with open(file_path, "rb") as fh:
            content = base64.b64encode(fh.read()).decode("ascii")
        payload = {"@odata.type": "#microsoft.graph.fileAttachment",
                   "name": os.path.basename(file_path), "contentBytes": content}
        a = await client.post(f"/me/messages/{draft_id}/attachments", json=payload)
        return {"attachment_id": a.get("id"), "name": a.get("name")}

    @mcp.tool()
    async def send_draft(draft_id: str) -> dict:
        """Send an existing draft."""
        await client.post(f"/me/messages/{draft_id}/send")
        return {"sent": draft_id}
```

- [ ] **Step 4: Run to verify it passes** → PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(tools): draft create/update/attach/send tools"
```

---

## Task 11: `tools/mail.py` (part 3) — organize + delete + gated permanent delete

**Files:**
- Modify: `src/outlook_personal_mcp/tools/mail.py` (fill `_register_organize`)
- Test: `tests/test_tools_mail_organize.py`

**Graph endpoints:** move `POST /me/messages/{id}/move {destinationId}`; copy `POST /me/messages/{id}/copy {destinationId}`; mark read/unread `PATCH /me/messages/{id} {isRead}`; flag `PATCH /me/messages/{id} {flag:{flagStatus}}`; delete `DELETE /me/messages/{id}` (→ Deleted Items); **permanent delete** `POST /me/messages/{id}/permanentDelete` — **VERIFY** against current Graph v1.0 docs at implementation time; if not on v1.0, fall back to `DELETE` then permanently remove from Deleted Items.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_mail_organize.py`

```python
import httpx, respx, pytest, json
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.config import Settings
from outlook_personal_mcp.tools import mail

class StubTokens:
    def get_token(self): return "T"

def _ctx(allow_perm=False):
    mcp = FastMCP("test")
    mail.register(mcp, GraphClient(StubTokens(), backoff_base=0),
                  Settings(client_id="abc", allow_permanent_delete=allow_perm))
    return mcp

@respx.mock
async def test_move_message(ctx_=None):
    ctx = _ctx()
    route = respx.post("https://graph.microsoft.com/v1.0/me/messages/m1/move").mock(
        return_value=httpx.Response(201, json={"id": "m1b"}))
    await ctx.call_tool("move_message", {"message_id": "m1", "destination_folder_id": "f2"})
    assert json.loads(route.calls.last.request.read())["destinationId"] == "f2"

@respx.mock
async def test_mark_read(ctx_=None):
    ctx = _ctx()
    route = respx.patch("https://graph.microsoft.com/v1.0/me/messages/m1").mock(
        return_value=httpx.Response(200, json={"id": "m1"}))
    await ctx.call_tool("mark_read", {"message_id": "m1", "read": True})
    assert json.loads(route.calls.last.request.read())["isRead"] is True

async def test_permanent_delete_absent_by_default():
    ctx = _ctx(allow_perm=False)
    names = {t.name for t in await ctx.list_tools()}
    assert "permanent_delete" not in names
    assert "delete_message" in names

async def test_permanent_delete_present_when_enabled():
    ctx = _ctx(allow_perm=True)
    names = {t.name for t in await ctx.list_tools()}
    assert "permanent_delete" in names
```

- [ ] **Step 2: Run to verify it fails** → FAIL.

- [ ] **Step 3: Implement `_register_organize` in `tools/mail.py`**

```python
def _register_organize(mcp, client, settings):
    @mcp.tool()
    async def move_message(message_id: str, destination_folder_id: str) -> dict:
        """Move a message to another folder."""
        m = await client.post(f"/me/messages/{message_id}/move",
                              json={"destinationId": destination_folder_id})
        return {"id": m.get("id"), "moved_to": destination_folder_id}

    @mcp.tool()
    async def copy_message(message_id: str, destination_folder_id: str) -> dict:
        """Copy a message to another folder."""
        m = await client.post(f"/me/messages/{message_id}/copy",
                              json={"destinationId": destination_folder_id})
        return {"id": m.get("id"), "copied_to": destination_folder_id}

    @mcp.tool()
    async def mark_read(message_id: str, read: bool = True) -> dict:
        """Mark a message read (read=True) or unread (read=False)."""
        await client.patch(f"/me/messages/{message_id}", json={"isRead": read})
        return {"id": message_id, "is_read": read}

    @mcp.tool()
    async def flag_message(message_id: str, flagged: bool = True) -> dict:
        """Flag or unflag a message."""
        status = "flagged" if flagged else "notFlagged"
        await client.patch(f"/me/messages/{message_id}", json={"flag": {"flagStatus": status}})
        return {"id": message_id, "flagged": flagged}

    @mcp.tool()
    async def delete_message(message_id: str) -> dict:
        """Delete a message (moves it to Deleted Items; reversible)."""
        await client.delete(f"/me/messages/{message_id}")
        return {"deleted": message_id, "permanent": False}

    if settings.allow_permanent_delete:
        @mcp.tool()
        async def permanent_delete(message_id: str) -> dict:
            """PERMANENTLY delete a message (irreversible). Enabled via env flag."""
            await client.post(f"/me/messages/{message_id}/permanentDelete")
            return {"deleted": message_id, "permanent": True}
```

- [ ] **Step 4: Run to verify it passes** → PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(tools): mail organize/delete + env-gated permanent delete"
```

---

## Task 12: `tools/calendar.py` (part 1) — list/get/search events + calendars

**Files:**
- Create (replace stub): `src/outlook_personal_mcp/tools/calendar.py`, `tests/test_tools_calendar_read.py`

**Graph endpoints:** calendars `GET /me/calendars`; events `GET /me/events?$top&$orderby=start/dateTime`; time window `GET /me/calendarView?startDateTime=&endDateTime=`; get `GET /me/events/{id}`; search `GET /me/events?$search="kql"`.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_calendar_read.py`

```python
import httpx, respx, pytest
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import calendar

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    calendar.register(mcp, GraphClient(StubTokens(), backoff_base=0))
    return mcp

@respx.mock
async def test_list_events_window_uses_calendarview(ctx):
    route = respx.get("https://graph.microsoft.com/v1.0/me/calendarView").mock(
        return_value=httpx.Response(200, json={"value": [
            {"id": "e1", "subject": "Sync",
             "start": {"dateTime": "2026-06-27T09:00:00", "timeZone": "UTC"},
             "end": {"dateTime": "2026-06-27T09:30:00", "timeZone": "UTC"}}]}))
    res = await ctx.call_tool("list_events",
                              {"start": "2026-06-27T00:00:00Z", "end": "2026-06-28T00:00:00Z"})
    assert route.calls.last.request.url.params["startDateTime"] == "2026-06-27T00:00:00Z"
    assert any("Sync" in str(c) for c in res[0])

@respx.mock
async def test_get_event(ctx):
    respx.get("https://graph.microsoft.com/v1.0/me/events/e1").mock(
        return_value=httpx.Response(200, json={"id": "e1", "subject": "1:1",
            "start": {"dateTime": "x", "timeZone": "UTC"},
            "end": {"dateTime": "y", "timeZone": "UTC"}}))
    res = await ctx.call_tool("get_event", {"event_id": "e1"})
    assert any("1:1" in str(c) for c in res[0])
```

- [ ] **Step 2: Run to verify it fails** → FAIL.

- [ ] **Step 3: Implement `tools/calendar.py` (read section; write section appended in Task 13)**

```python
from __future__ import annotations
from ..models import shape_event

_SELECT = "id,subject,start,end,location,isAllDay,organizer"

def register(mcp, client):
    @mcp.tool()
    async def list_calendars() -> list[dict]:
        """List the user's calendars."""
        data = await client.get("/me/calendars", params={"$select": "id,name,isDefaultCalendar"})
        return [{"id": c["id"], "name": c.get("name"),
                 "is_default": c.get("isDefaultCalendar")} for c in data.get("value", [])]

    @mcp.tool()
    async def list_events(start: str | None = None, end: str | None = None,
                          top: int = 25) -> list[dict]:
        """List events. If start and end (ISO 8601) are given, returns that window via calendarView."""
        if start and end:
            data = await client.get("/me/calendarView",
                                    params={"startDateTime": start, "endDateTime": end,
                                            "$select": _SELECT, "$top": top,
                                            "$orderby": "start/dateTime"})
        else:
            data = await client.get("/me/events",
                                    params={"$select": _SELECT, "$top": top,
                                            "$orderby": "start/dateTime"})
        return [shape_event(e) for e in data.get("value", [])]

    @mcp.tool()
    async def search_events(query: str, top: int = 25) -> list[dict]:
        """Search events by free text."""
        data = await client.get("/me/events",
                                params={"$search": f'"{query}"', "$select": _SELECT, "$top": top})
        return [shape_event(e) for e in data.get("value", [])]

    @mcp.tool()
    async def get_event(event_id: str) -> dict:
        """Get one event including body and attendees."""
        e = await client.get(f"/me/events/{event_id}",
                             params={"$select": _SELECT + ",body,attendees,onlineMeeting"})
        out = shape_event(e)
        out["body"] = (e.get("body") or {}).get("content")
        out["attendees"] = [a["emailAddress"]["address"] for a in e.get("attendees", [])]
        join = (e.get("onlineMeeting") or {}).get("joinUrl")
        if join:
            out["join_url"] = join
        return out

    _register_calendar_write(mcp, client)
```

> As in Task 8, add a bottom-of-module `def _register_calendar_write(*a): pass` placeholder so Task 12 is green, filled in Task 13.

- [ ] **Step 4: Run to verify it passes** → PASS.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat(tools): calendar read tools (calendars/list/search/get)"
```

---

## Task 13: `tools/calendar.py` (part 2) — create/update/delete/respond/availability + final wiring

**Files:**
- Modify: `src/outlook_personal_mcp/tools/calendar.py` (fill `_register_calendar_write`)
- Modify: `tests/test_server.py` (re-add the broader registration asserts from Task 6 Step 7)
- Test: `tests/test_tools_calendar_write.py`

**Graph endpoints:** create `POST /me/events {event}`; update `PATCH /me/events/{id}`; delete `DELETE /me/events/{id}`; respond `POST /me/events/{id}/{accept|decline|tentativelyAccept} {comment, sendResponse}`; availability `POST /me/calendar/getSchedule {schedules, startTime, endTime, availabilityViewInterval}`.

- [ ] **Step 1: Write the failing test** — `tests/test_tools_calendar_write.py`

```python
import httpx, respx, pytest, json
from mcp.server.fastmcp import FastMCP
from outlook_personal_mcp.graph import GraphClient
from outlook_personal_mcp.tools import calendar

class StubTokens:
    def get_token(self): return "T"

@pytest.fixture
def ctx():
    mcp = FastMCP("test")
    calendar.register(mcp, GraphClient(StubTokens(), backoff_base=0))
    return mcp

@respx.mock
async def test_create_event_builds_payload(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/events").mock(
        return_value=httpx.Response(201, json={"id": "e9", "subject": "Demo",
            "start": {"dateTime": "2026-07-01T15:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-07-01T15:30:00", "timeZone": "UTC"}}))
    await ctx.call_tool("create_event", {"subject": "Demo",
        "start": "2026-07-01T15:00:00", "end": "2026-07-01T15:30:00",
        "attendees": ["a@x.com"], "is_online_meeting": True})
    body = json.loads(route.calls.last.request.read())
    assert body["subject"] == "Demo"
    assert body["isOnlineMeeting"] is True
    assert body["attendees"][0]["emailAddress"]["address"] == "a@x.com"

@respx.mock
async def test_respond_event_accept(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/events/e1/accept").mock(
        return_value=httpx.Response(202))
    await ctx.call_tool("respond_event", {"event_id": "e1", "response": "accept", "comment": "yes"})
    body = json.loads(route.calls.last.request.read())
    assert body["sendResponse"] is True and body["comment"] == "yes"

@respx.mock
async def test_find_availability(ctx):
    route = respx.post("https://graph.microsoft.com/v1.0/me/calendar/getSchedule").mock(
        return_value=httpx.Response(200, json={"value": [{"scheduleId": "a@x.com",
            "availabilityView": "000"}]}))
    res = await ctx.call_tool("find_availability", {"emails": ["a@x.com"],
        "start": "2026-07-01T09:00:00", "end": "2026-07-01T17:00:00"})
    assert json.loads(route.calls.last.request.read())["schedules"] == ["a@x.com"]
    assert any("a@x.com" in str(c) for c in res[0])
```

- [ ] **Step 2: Run to verify it fails** → FAIL.

- [ ] **Step 3: Implement `_register_calendar_write` in `tools/calendar.py`**

```python
from ..models import EventInput, shape_event

_RESPONSE_ACTIONS = {"accept": "accept", "decline": "decline", "tentative": "tentativelyAccept"}

def _register_calendar_write(mcp, client):
    @mcp.tool()
    async def create_event(subject: str, start: str, end: str, time_zone: str = "UTC",
                           attendees: list[str] | None = None, location: str | None = None,
                           body: str | None = None, is_online_meeting: bool = False) -> dict:
        """Create a calendar event. Times are ISO 8601 in the given time_zone."""
        payload = EventInput(subject=subject, start=start, end=end, time_zone=time_zone,
                             attendees=attendees or [], location=location, body=body,
                             is_online_meeting=is_online_meeting).to_graph()
        e = await client.post("/me/events", json=payload)
        return shape_event(e)

    @mcp.tool()
    async def update_event(event_id: str, subject: str | None = None,
                           start: str | None = None, end: str | None = None,
                           time_zone: str = "UTC", location: str | None = None) -> dict:
        """Update fields on an existing event (only provided fields change)."""
        patch: dict = {}
        if subject is not None:
            patch["subject"] = subject
        if start is not None:
            patch["start"] = {"dateTime": start, "timeZone": time_zone}
        if end is not None:
            patch["end"] = {"dateTime": end, "timeZone": time_zone}
        if location is not None:
            patch["location"] = {"displayName": location}
        e = await client.patch(f"/me/events/{event_id}", json=patch)
        return shape_event(e)

    @mcp.tool()
    async def delete_event(event_id: str) -> dict:
        """Delete/cancel a calendar event."""
        await client.delete(f"/me/events/{event_id}")
        return {"deleted": event_id}

    @mcp.tool()
    async def respond_event(event_id: str, response: str, comment: str = "",
                            send_response: bool = True) -> dict:
        """Respond to a meeting invite. response is one of: accept, decline, tentative."""
        action = _RESPONSE_ACTIONS.get(response)
        if not action:
            raise ValueError("response must be accept, decline, or tentative")
        await client.post(f"/me/events/{event_id}/{action}",
                          json={"comment": comment, "sendResponse": send_response})
        return {"event_id": event_id, "response": response}

    @mcp.tool()
    async def find_availability(emails: list[str], start: str, end: str,
                                time_zone: str = "UTC", interval_minutes: int = 30) -> list[dict]:
        """Free/busy across the given people for a time window (Graph getSchedule)."""
        payload = {"schedules": emails,
                   "startTime": {"dateTime": start, "timeZone": time_zone},
                   "endTime": {"dateTime": end, "timeZone": time_zone},
                   "availabilityViewInterval": interval_minutes}
        data = await client.post("/me/calendar/getSchedule", json=payload)
        return [{"email": s.get("scheduleId"), "availability_view": s.get("availabilityView")}
                for s in data.get("value", [])]
```

- [ ] **Step 4: Re-strengthen `tests/test_server.py`** — restore the asserts deferred in Task 6:

```python
    assert "send_mail" in names
    assert "list_events" in names
    assert "create_event" in names
```

- [ ] **Step 5: Run the full suite** → `uv run pytest -q` → Expected: ALL PASS.

- [ ] **Step 6: Lint** → `uv run ruff check .` → Expected: clean.

- [ ] **Step 7: Commit**

```bash
git add -A && git commit -m "feat(tools): calendar create/update/delete/respond/availability + full wiring"
```

---

## Task 14: README, `.env.example`, CI

**Files:**
- Create: `README.md`, `.github/workflows/ci.yml`

- [ ] **Step 1: Write `README.md`** with these sections (complete prose, no placeholders beyond `<your-app-client-id>`):
  1. What it is (personal Outlook.com mailbox + calendar over MCP; Claude Code + Codex).
  2. **Azure app registration** — the 5 steps from DESIGN §7 (Personal Microsoft accounts only; Allow public client flows = Yes; delegated `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`; copy client id).
  3. **Install & login** — `uvx --from git+https://github.com/salahawad/outlook-personal-mcp outlook-personal-mcp login` then approve the device code.
  4. **Claude Code config** — the `.mcp.json` block from DESIGN §7.
  5. **Codex config** — the `~/.codex/config.toml` block from DESIGN §7.
  6. **Tools list** — table of every tool + one-line description.
  7. **Permanent delete** — explain the `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE` gate and the irreversibility warning.
  8. **Security** — token cache location/perms; never commit the cache; revoke by deleting the Azure app or clearing the cache.
  9. License (MIT).

- [ ] **Step 2: Write `.github/workflows/ci.yml`**

```yaml
name: ci
on:
  push:
  pull_request:
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: astral-sh/setup-uv@v3
      - run: uv python install 3.11
      - run: uv pip install --system -e ".[dev]"
      - run: ruff check .
      - run: pytest -q
```

- [ ] **Step 3: Run the full gate locally** → `uv run ruff check . && uv run pytest -q` → Expected: clean + all pass.

- [ ] **Step 4: Commit**

```bash
git add -A && git commit -m "docs: README (Azure setup + Claude Code/Codex config) and CI workflow"
```

---

## Task 15: Manual smoke test + publish to GitHub (requires user confirmation)

**Files:** none (operational).

- [ ] **Step 1: Real login** — `OUTLOOK_MCP_CLIENT_ID=<id> uv run outlook-personal-mcp login`, approve device code, confirm "Login successful".

- [ ] **Step 2: Smoke a read tool via MCP Inspector or a tiny client** — call `whoami` and `list_messages` against the live account; confirm real data returns. (Document the exact command used.)

- [ ] **Step 3: Smoke `list_events`** for today's window; confirm events return.

- [ ] **Step 4: Create the public GitHub repo** — **CONFIRM WITH USER FIRST.**

```bash
gh repo create salahawad/outlook-personal-mcp --public --source=. --remote=origin --push
```

- [ ] **Step 5: Verify CI goes green** on GitHub Actions; fix any environment-specific failures.

---

## Self-Review (completed by plan author)

**Spec coverage** — every DESIGN section maps to a task: config §7→T2; auth §4→T3; graph §6→T4; models→T5; server/CLI→T6; folders→T7; mail read→T8; send→T9; drafts→T10; organize/delete/perm-delete gate §5→T11; calendar read→T12; calendar write + availability→T13; README/Azure/Claude+Codex config §7 + CI §8→T14; smoke + publish→T15. Calendar scope (added post-spec) covered in T12–T13 and scopes set in T2. No gaps.

**Placeholder scan** — no TBD/TODO. The only literal placeholder is `<your-app-client-id>` / `<id>`, which is intentional user-supplied input. Two endpoints are flagged **VERIFY** (`permanentDelete`, with an explicit fallback) — these are correctness checks, not missing content.

**Type consistency** — `GraphClient` methods (`get/post/patch/delete/request`, `raw=` kwarg) used consistently T4→T13; `register(mcp, client)` for folders/drafts/calendar and `register(mcp, client, settings)` for mail are consistent T6→T13; `shape_message`/`shape_event`/`SendMailInput.to_graph`/`EventInput.to_graph` signatures match between T5 and their consumers; the `_register_write`/`_register_organize`/`_register_calendar_write` placeholder→fill pattern is called out in T8/T12 and filled in T9/T11/T13.
