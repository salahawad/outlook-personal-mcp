<!-- mcp-name: io.github.salahawad/outlook-personal-mcp -->
# outlook-personal-mcp

An MCP server that gives Claude Code and Codex full control of a personal Outlook.com mailbox and calendar via the Microsoft Graph API. Written in Python, speaks the MCP stdio transport, and uses per-user device-code OAuth so your credentials never leave your machine. MIT licensed.

---

## Features

- **Mail** — list, search, read, send, reply, forward, move, copy, flag, mark read/unread, delete (soft or hard)
- **Drafts** — create, update, attach local files, send
- **Folders** — list, create, rename, delete
- **Calendar** — list calendars, list/search/get/create/update/delete events, respond to invites (accept/decline/tentative), check free/busy availability
- **Per-user OAuth** — you register your own free Azure app; the server authenticates with your Microsoft account and caches the token locally
- **Local stdio** — runs as a child process of the MCP host; your mailbox data never transits a third party

---

## Prerequisites

- Python 3.10 or later
- [`uv`](https://docs.astral.sh/uv/) (fast Python package and tool runner)
- A personal Microsoft account (Outlook.com, Hotmail, Live, etc.)
- A free Azure app registration (see below — takes about three minutes)

---

## Azure App Registration

1. Go to [https://portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it anything you like (e.g. `outlook-personal-mcp`). Under **Supported account types** choose **Personal Microsoft accounts only**. No redirect URI is needed. Click **Register**.
3. Open the app → **Authentication** → **Advanced settings** → **Allow public client flows** → set to **Yes** → **Save**. (This is required for the device-code login flow used by this server.)
4. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, then add:
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`

   (`User.Read` is included by default; `offline_access` is requested automatically at runtime — you do not need to add it.)
5. Copy the **Application (client) ID** from the **Overview** page. This is your `OUTLOOK_MCP_CLIENT_ID`.

---

## Install & First-Time Login

Run the one-time interactive login. It prints a short URL and a code; open the URL in any browser, enter the code, approve the permissions, and you are done. The token is cached at `~/.config/outlook-personal-mcp/token_cache.bin` (mode 600) and refreshed automatically on subsequent runs — you will not be prompted again unless the refresh token expires or is revoked.

```bash
OUTLOOK_MCP_CLIENT_ID=<your-app-client-id> \
  uvx --from git+https://github.com/salahawad/outlook-personal-mcp outlook-personal-mcp login
```

---

## Configure Claude Code

Add the server to your project's `.mcp.json` (or `~/.claude/.mcp.json` for all projects):

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

Alternatively, use the CLI: `claude mcp add`.

---

## Configure Codex

Add the server to `~/.codex/config.toml`:

```toml
[mcp_servers.outlook]
command = "uvx"
args = ["--from", "git+https://github.com/salahawad/outlook-personal-mcp", "outlook-personal-mcp"]
env = { OUTLOOK_MCP_CLIENT_ID = "<your-app-client-id>" }
```

---

## Configuration (Environment Variables)

| Variable | Required | Default | Description |
|---|---|---|---|
| `OUTLOOK_MCP_CLIENT_ID` | **Yes** | — | Azure app's Application (client) ID |
| `OUTLOOK_MCP_AUTHORITY` | No | `https://login.microsoftonline.com/consumers` | MSAL authority URL (change only if you move to a work/school tenant) |
| `OUTLOOK_MCP_TOKEN_CACHE` | No | `~/.config/outlook-personal-mcp/token_cache.bin` | Path to the MSAL token cache file |
| `OUTLOOK_MCP_FILE_ROOT` | No | `~/.local/share/outlook-personal-mcp/files` | Only files under this directory can be read by `add_attachment` or written by `download_attachment` |
| `OUTLOOK_MCP_MAX_FILE_BYTES` | No | `3145728` | Maximum bytes allowed for local attachment reads and attachment downloads |
| `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE` | No | `false` | Set to `true` to enable the `permanent_delete` tool (irreversible — see below) |
| `OUTLOOK_MCP_DEBUG` | No | `false` | Set to `true` to log each Graph request's method, URL, and HTTP status code to stderr. Never logs tokens or message content. |

---

## Tools

### Account

| Tool | Description |
|---|---|
| `whoami` | Return the signed-in user's Microsoft account profile |

### Mail

| Tool | Description |
|---|---|
| `list_messages` | List messages (newest first); optionally filter by folder or unread-only |
| `search_messages` | Full-text search across the entire mailbox (Graph `$search`) |
| `get_message` | Get a single message; optionally include the full body |
| `list_attachments` | List a message's attachments (id, name, size, content type) |
| `download_attachment` | Download an attachment to a path under `OUTLOOK_MCP_FILE_ROOT` (refuses to overwrite an existing file) |
| `send_mail` | Send an email |
| `reply` | Reply to a message (`reply_all` to reply to everyone) |
| `forward` | Forward a message to recipients with an optional comment |
| `move_message` | Move a message to another folder |
| `copy_message` | Copy a message to another folder |
| `mark_read` | Mark a message read or unread |
| `flag_message` | Flag or unflag a message |
| `delete_message` | Delete a message (moves it to Deleted Items; reversible) |
| `permanent_delete` | **Permanently** delete a message (irreversible). Only available when `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE=true` |

### Folders

| Tool | Description |
|---|---|
| `list_folders` | List mail folders with unread and total message counts |
| `create_folder` | Create a mail folder, optionally nested under a parent |
| `rename_folder` | Rename a mail folder |
| `delete_folder` | Delete a mail folder (moves it to Deleted Items) |

### Drafts

| Tool | Description |
|---|---|
| `create_draft` | Create a draft message (not sent) |
| `update_draft` | Update a draft's subject and/or body |
| `add_attachment` | Attach a local file (a regular, non-symlink file under `OUTLOOK_MCP_FILE_ROOT`) to a draft |
| `send_draft` | Send an existing draft |

### Calendar

| Tool | Description |
|---|---|
| `list_calendars` | List the user's calendars |
| `list_events` | List events; if `start`/`end` (ISO 8601) are given, returns that time window |
| `search_events` | Search events by free text |
| `get_event` | Get one event including body, attendees, and online meeting link |
| `create_event` | Create a calendar event with optional attendees and online meeting |
| `update_event` | Update fields on an existing event (only provided fields change) |
| `delete_event` | Delete/cancel a calendar event |
| `respond_event` | Respond to a meeting invite: `accept`, `decline`, or `tentative` |
| `find_availability` | Get free/busy availability for a list of people over a time window |

---

## Permanent Delete

The `permanent_delete` tool bypasses the Deleted Items folder and removes a message irreversibly. It is **disabled by default** — when `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE` is not set (or is `false`), the tool is not registered with the MCP server at all and will not appear in the tool list.

To enable it, set `OUTLOOK_MCP_ALLOW_PERMANENT_DELETE=true` in the server's environment block in your `.mcp.json` / `config.toml`. Only do this if you understand the consequences: there is no undo and no Recoverable Items path for personal accounts.

---

## Security

- **Token cache is a credential.** The file at `~/.config/outlook-personal-mcp/token_cache.bin` contains a long-lived refresh token. It is written with mode `600`, but treat it like a password — never commit it, never share it, and store it on an encrypted volume.
- **Data stays local.** The server runs as a child process of Claude Code / Codex over stdio. Your mailbox content is passed directly between the MCP host and the Microsoft Graph API; no third-party relay is involved.
- **Revocation.** To revoke access, delete the token cache file and/or navigate to [https://account.microsoft.com/permissions](https://account.microsoft.com/permissions) to remove the Azure app's consent. You can also delete the Azure app registration entirely from the portal.
- **File paths.** `download_attachment` writes only under `OUTLOOK_MCP_FILE_ROOT` and refuses to overwrite existing files. `add_attachment` reads only regular, non-symlink files under `OUTLOOK_MCP_FILE_ROOT`. Relative paths are resolved under that root; absolute paths outside it (and any path traversing a symlink) are rejected. Both tools enforce `OUTLOOK_MCP_MAX_FILE_BYTES`. Review these paths before confirming any tool call that touches the filesystem.

---

## Development

```bash
git clone https://github.com/salahawad/outlook-personal-mcp
cd outlook-personal-mcp
uv venv && uv pip install -e ".[dev]"
uv run pytest
uv run ruff check .
```

---

## Privacy

This server runs entirely on your machine and sends data only between your machine and Microsoft's Graph API — no third-party relay, no telemetry, and the maintainer receives nothing. OAuth tokens are cached locally at `~/.config/outlook-personal-mcp/token_cache.bin` (mode 600). See [PRIVACY.md](PRIVACY.md) for the full privacy policy.

---

## License

MIT — see [LICENSE](LICENSE).
