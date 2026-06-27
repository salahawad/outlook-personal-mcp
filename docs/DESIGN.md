# outlook-personal-mcp — Design

> A generic, open-source **Model Context Protocol (MCP) server** that gives an AI
> coding agent (Claude Code, Codex, or any stdio-MCP host) full control of a
> **personal Microsoft / Outlook.com mailbox and calendar** via the Microsoft
> Graph API.
>
> Status: **approved design** (2026-06-27). Implementation pending.

## 1. Goals & non-goals

**Goals**

- Let any user connect their own personal Microsoft account (`@outlook.com`,
  `@hotmail.com`, `@live.com`) to an AI agent over MCP.
- **Full mailbox control**: read, search, send, reply/forward, drafts, folders,
  flags, move/copy, delete — including (guarded) permanent delete.
- **Full calendar control**: list/search events, create/update/delete events,
  accept/decline/tentatively-accept invites, and find free/busy availability.
- Zero-friction for the *operator* of the AI agent: install with `uvx`, run, do a
  one-time device-code login, done.
- Generic & reusable: nothing hardcoded to one account. Each user supplies their
  own Azure app client id and signs in as themselves.
- Works identically for **Claude Code** and **Codex** (both speak stdio MCP).

**Non-goals (v1)**

- Contacts, tasks, OneDrive — documented as future extensions, not in v1.
- Work/school (Entra ID org) accounts — the authority defaults to `consumers`
  (personal accounts). Org support is a config override, not a tested path in v1.
- A hosted/remote multi-tenant server. v1 is **local stdio**, one user per process.

## 2. High-level architecture

```
┌────────────────┐   stdio (MCP)   ┌──────────────────────────────────────────┐
│ Claude Code /  │ ───────────────▶│           outlook-personal-mcp           │
│ Codex (host)   │                 │  ┌────────┐  ┌────────┐  ┌─────────────┐  │
└────────────────┘                 │  │server  │─▶│ tools/ │─▶│  graph.py   │──┼──▶ Microsoft
                                   │  │(FastMCP)│  │ mail   │  │ httpx REST  │  │    Graph API
                                   │  └────────┘  │ folders│  │ +retry/429  │  │  (graph.microsoft.com)
                                   │       │      │ drafts │  └─────┬───────┘  │
                                   │       │      └────────┘        │          │
                                   │   ┌───▼────┐             ┌─────▼───────┐  │
                                   │   │config  │             │  auth.py    │──┼──▶ login.microsoftonline.com
                                   │   │(env)   │             │ MSAL device │  │    /consumers (device code)
                                   │   └────────┘             │ code + cache│  │
                                   │                          └─────────────┘  │
                                   └──────────────────────────────────────────┘
                                                   token_cache.bin (chmod 600)
```

Transport is **stdio**: each user runs one process with their own credentials. No
inbound network listener, no shared state.

## 3. Module breakdown

Each module is single-purpose and independently testable.

| Module | Responsibility | Depends on |
| --- | --- | --- |
| `config.py` | Load + validate env settings into a frozen `Settings` object. | env only |
| `auth.py` | MSAL `PublicClientApplication`; device-code login; persistent, file-backed token cache; silent refresh; expose `get_access_token()`. | `config`, `msal` |
| `graph.py` | Thin async `httpx` client to Graph REST: inject bearer, JSON helpers, throttling/5xx retry honoring `Retry-After`, map Graph errors → typed exceptions. | `auth`, `httpx` |
| `models.py` | Pydantic schemas for tool inputs/outputs (validation + clean MCP schemas). | `pydantic` |
| `tools/mail.py` | Read/search/send/reply/forward/move/flag/delete message tools. | `graph`, `models` |
| `tools/folders.py` | List/create/rename/delete mail folders. | `graph`, `models` |
| `tools/drafts.py` | Create/update/send drafts; attachments. | `graph`, `models` |
| `tools/calendar.py` | List/search/create/update/delete events; respond to invites; free/busy. | `graph`, `models` |
| `server.py` | FastMCP app; register tools; stdio entrypoint (`__main__`). | all of the above |

**Design rule:** tool modules never touch MSAL or httpx directly — they call
`graph.py`. `graph.py` never knows about MCP. This keeps each layer swappable and
mockable.

## 4. Authentication design

- **Flow:** OAuth2 **device authorization grant** (public client, no secret).
  Ideal for headless CLI hosts — no localhost redirect server needed.
- **Per-user app:** the user registers their own Azure app and provides its client
  id via `OUTLOOK_MCP_CLIENT_ID`. Each install is self-owned and isolated.
- **Authority:** default `https://login.microsoftonline.com/consumers` (personal
  accounts). Overridable via `OUTLOOK_MCP_AUTHORITY` (e.g. `common`).
- **Scopes:** `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`,
  `offline_access`.
- **First run:** if no valid cached token, the relevant tool call returns a clear
  message containing the verification URL + user code. After the user approves in a
  browser, tokens are cached and the call can be retried. (A `login` tool and a
  `--login` CLI subcommand let users pre-authenticate explicitly.)
- **Token cache:** MSAL `SerializableTokenCache` persisted to
  `~/.config/outlook-personal-mcp/token_cache.bin`, created `chmod 600`. Path
  overridable via `OUTLOOK_MCP_TOKEN_CACHE`. Silent refresh on every call;
  re-prompt only when the refresh token is gone/expired.

## 5. Tool catalog (v1 — full mailbox + calendar)

Identity: `whoami`.

Folders: `list_folders`, `create_folder`, `rename_folder`, `delete_folder`.

Read: `list_messages` (folder + filter/top/skip), `search_messages` (Graph
`$search`), `get_message` (optional full body + attachment metadata),
`download_attachment`.

Compose/send: `send_mail` (to/cc/bcc, subject, html|text body, attachments),
`reply`, `reply_all`, `forward`.

Drafts: `create_draft`, `update_draft`, `send_draft`.

Organize: `move_message`, `copy_message`, `mark_read`, `mark_unread`,
`flag_message`.

Delete: `delete_message` (→ Deleted Items, reversible). **`permanent_delete`**
(hard delete) is registered **only when** `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE=true`
(default `false`) — full power without accidental data loss.

Calendar: `list_calendars`; `list_events` (calendar + time window/filter),
`search_events`, `get_event`; `create_event` (subject, start/end + time zone,
attendees, location, body, online-meeting flag), `update_event`, `delete_event`;
`respond_event` (`accept` | `decline` | `tentative`, optional message);
`find_availability` (free/busy across given times, via Graph `getSchedule`). The
calendar delete tool removes the event (cancellation notices follow Graph's normal
behavior for organizer vs attendee); it is **not** gated behind the permanent-delete
flag since calendar deletes are not destructive to mail data.

Each tool has a pydantic input model; outputs are compact JSON (ids, subjects,
from/to, received time, preview) rather than raw Graph payloads, to keep token use
low. `get_message` can return full content on request.

## 6. Error handling

- **401 Unauthorized** → attempt MSAL silent refresh; if that fails, return an
  actionable "re-authenticate" message with the device-code URL.
- **403 Forbidden** → message naming the missing scope and how to consent.
- **429 / 503** → respect `Retry-After`; capped exponential backoff (e.g. max 3
  retries); then surface a throttling error.
- **404** → clear "message/folder not found" with the offending id.
- All tool errors are returned as structured MCP errors, never stack traces.
- **Security:** access/refresh tokens and `Authorization` headers are never logged;
  message bodies are not logged by default (opt-in `OUTLOOK_MCP_DEBUG`).

## 7. Configuration

Environment variables (read by `config.py`):

| Var | Required | Default | Purpose |
| --- | --- | --- | --- |
| `OUTLOOK_MCP_CLIENT_ID` | **yes** | — | Azure app (public client) id. |
| `OUTLOOK_MCP_AUTHORITY` | no | `…/consumers` | Token authority. |
| `OUTLOOK_MCP_TOKEN_CACHE` | no | `~/.config/outlook-personal-mcp/token_cache.bin` | Token cache path. |
| `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE` | no | `false` | Register the hard-delete tool. |
| `OUTLOOK_MCP_DEBUG` | no | `false` | Verbose (still redacts secrets). |

### Host config snippets (shipped in README)

**Claude Code** (`.mcp.json` or `claude mcp add`):
```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/salahawad/outlook-personal-mcp", "outlook-personal-mcp"],
      "env": { "OUTLOOK_MCP_CLIENT_ID": "<your-app-client-id>" }
    }
  }
}
```

**Codex** (`~/.codex/config.toml`):
```toml
[mcp_servers.outlook]
command = "uvx"
args = ["--from", "git+https://github.com/salahawad/outlook-personal-mcp", "outlook-personal-mcp"]
env = { OUTLOOK_MCP_CLIENT_ID = "<your-app-client-id>" }
```

### Azure app registration (README, step-by-step)

1. Azure Portal → App registrations → New registration.
2. Supported account types: **Personal Microsoft accounts only**.
3. No redirect URI needed (device code). Under **Authentication**, set **Allow
   public client flows = Yes**.
4. API permissions → Microsoft Graph → **Delegated** → add `Mail.ReadWrite`,
   `Mail.Send`, `Calendars.ReadWrite` (`User.Read` is added by default).
   `offline_access` is requested at runtime.
5. Copy the **Application (client) ID** → that's `OUTLOOK_MCP_CLIENT_ID`.

## 8. Testing & CI

- **pytest** with Graph HTTP fully mocked (`respx`/`httpx.MockTransport`) — no live
  calls in CI.
- Coverage targets: token-cache load/save + silent-refresh logic; `graph.py` retry
  on 429 honoring `Retry-After`; error mapping; each tool's request shaping and
  response parsing; the permanent-delete env gate (registered iff flag set).
- **ruff** lint. **GitHub Actions** runs ruff + pytest on push/PR.

## 9. Repository layout

```
outlook-personal-mcp/
├── pyproject.toml            # build (hatch/uv), deps, console script `outlook-personal-mcp`
├── README.md                 # quickstart, Azure setup, Claude Code + Codex config
├── LICENSE                   # MIT
├── .env.example
├── .github/workflows/ci.yml
├── docs/DESIGN.md            # this document
├── src/outlook_personal_mcp/
│   ├── __init__.py
│   ├── __main__.py           # `python -m outlook_personal_mcp`
│   ├── server.py
│   ├── auth.py
│   ├── graph.py
│   ├── config.py
│   ├── models.py
│   └── tools/{__init__,mail,folders,drafts,calendar}.py
└── tests/
    ├── conftest.py
    ├── test_auth.py
    ├── test_graph.py
    └── test_tools_{mail,folders,drafts,calendar}.py
```

## 10. Dependencies

`mcp` (FastMCP), `msal`, `httpx`, `pydantic`. Dev: `pytest`, `respx`, `ruff`.

## 11. Future work (out of v1 scope)

Contacts, tasks, OneDrive attachments, shared/other-user calendars, PyPI
publishing, optional org-account profile, optional `auth-code + PKCE` flow for GUI
hosts.

## 12. License

MIT — generic, permissive, suitable for broad public reuse.
