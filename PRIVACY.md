# Privacy Policy — outlook-personal-mcp

**Last updated: 2026-06-27**

## What this server is

`outlook-personal-mcp` is a local, open-source MCP server that you download, install, and run yourself on your own machine. It is not a hosted service and has no backend operated by the maintainer.

Source code: [https://github.com/salahawad/outlook-personal-mcp](https://github.com/salahawad/outlook-personal-mcp)

---

## Data we collect: none

The maintainer of this project receives no data from you whatsoever. There is no telemetry, no analytics, no crash reporting, no logging to a remote server, and no account creation. Nothing you do while using this server is ever transmitted to the maintainer or any third party other than Microsoft (see below).

---

## How your data flows

When you use this server:

1. **Your machine → Microsoft Graph API.** Tool calls (reading mail, sending mail, managing calendar events, checking availability) are fulfilled by making HTTPS requests directly from your machine to `https://graph.microsoft.com`. No intermediate relay is involved. The data path is your MCP host (Claude Code, Codex) → this server process (local stdio) → Microsoft's servers.

2. **Your machine → your MCP host.** The server returns mail and calendar content over stdio to your MCP host (e.g., Claude Code running locally). That data stays on your machine and is governed by the privacy policy of your MCP host.

3. **Nowhere else.** No other party receives your data.

---

## OAuth tokens and credentials

- Authentication uses Microsoft's device-code OAuth flow with your **own** Azure app registration (Application ID). You create this app in your own Azure account; the maintainer never has access to it.
- After you complete the device-code login, MSAL caches a refresh token at `~/.config/outlook-personal-mcp/token_cache.bin`. This file is written with mode `600` (readable only by you).
- The token cache contains a long-lived Microsoft refresh token and should be treated as a credential. Do not commit it, do not share it, and consider storing it on an encrypted volume.
- To revoke access: delete `~/.config/outlook-personal-mcp/token_cache.bin` and/or visit [https://account.microsoft.com/permissions](https://account.microsoft.com/permissions) to remove the Azure app's consent. You can also delete the Azure app registration entirely from the Azure portal.

---

## Local file access

- The `download_attachment` tool writes an email attachment to a local file path that you or your AI agent provides.
- The `add_attachment` tool reads a local file from a path that you or your AI agent provides.
- The server does not scan, index, or transmit any other files on your system.

---

## Mail and calendar content

Mail bodies, subjects, sender addresses, calendar event details, and attendee information are fetched from Microsoft Graph and returned to your MCP host to fulfill tool calls. This data is processed in memory only. The server does not write it to disk, log it, or transmit it anywhere other than back to your MCP host over stdio.

---

## Microsoft's privacy policy

Because this server accesses the Microsoft Graph API on your behalf, Microsoft's own privacy practices apply to that data. See Microsoft's Privacy Statement: [https://privacy.microsoft.com/privacystatement](https://privacy.microsoft.com/privacystatement)

---

## Contact

Questions or concerns? Open an issue at: [https://github.com/salahawad/outlook-personal-mcp/issues](https://github.com/salahawad/outlook-personal-mcp/issues)
