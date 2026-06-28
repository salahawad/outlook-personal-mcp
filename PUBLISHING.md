# Publishing Runbook

This document covers everything needed to publish `outlook-personal-mcp` to PyPI, the MCP Registry, and as an MCPB desktop extension.

---

## 1. PyPI Trusted Publisher Setup (one-time)

PyPI Trusted Publishing lets the GitHub Actions workflow publish without storing a PyPI API token as a secret. It uses OIDC identity tokens, which are issued automatically during the workflow run.

**Steps on pypi.org:**

1. Log in to [https://pypi.org](https://pypi.org) and go to your account's **Your projects** page.
2. Click **Manage** next to `mcp-outlook-personal`.
3. Select **Publishing** in the left sidebar.
4. Under **Add a new publisher**, fill in:
   - **Publisher:** GitHub Actions
   - **Owner:** `salahawad`
   - **Repository name:** `outlook-personal-mcp`
   - **Workflow name:** `publish.yml`
   - **Environment name:** `pypi`
5. Click **Add**.

**Steps on GitHub:**

1. In the repository, go to **Settings → Environments → New environment**.
2. Name it exactly `pypi`.
3. (Recommended) Add protection rules: require a reviewer, or restrict to tag patterns like `v*`.

No secrets need to be added. The `permissions: id-token: write` in the workflow is what allows GitHub to issue the OIDC token.

---

## 2. Release to PyPI

Once the trusted publisher is configured:

```bash
# Tag the release locally
git tag v0.1.0
git push origin v0.1.0
```

Alternatively, create a GitHub Release via the web UI and publish it. Either trigger fires `publish.yml`, which:

1. Checks out the repo and sets up `uv`.
2. Installs dev dependencies, runs `ruff`, and runs the pytest suite.
3. Runs `uv build` to produce `dist/outlook_personal_mcp-<version>-py3-none-any.whl` and `dist/outlook_personal_mcp-<version>.tar.gz`.
4. Publishes both to PyPI via `pypa/gh-action-pypi-publish@release/v1` using OIDC (no token required).

The `README.md` (referenced as `readme = "README.md"` in `pyproject.toml`) becomes the PyPI long-description automatically. It contains the `<!-- mcp-name: ... -->` marker that the MCP Registry needs for validation.

**Version consistency check** — before tagging, make sure these three values match:

| Location | Field | Value |
|---|---|---|
| `pyproject.toml` | `version` | `0.1.0` |
| `server.json` | `version` and `packages[0].version` | `0.1.0` |
| Git tag / GitHub release | tag name | `v0.1.0` |

---

## 3. Publish to the Official MCP Registry

The MCP Registry ([registry.modelcontextprotocol.io](https://registry.modelcontextprotocol.io)) hosts metadata (not artifacts). The package must already be on PyPI before you publish to the registry.

**Install `mcp-publisher`:**

```bash
# macOS/Linux — downloads the latest binary from GitHub releases
curl -L "https://github.com/modelcontextprotocol/registry/releases/latest/download/mcp-publisher_$(uname -s | tr '[:upper:]' '[:lower:]')_$(uname -m | sed 's/x86_64/amd64/;s/aarch64/arm64/').tar.gz" \
  | tar xz mcp-publisher && sudo mv mcp-publisher /usr/local/bin/

# macOS with Homebrew
brew install mcp-publisher

# Verify
mcp-publisher --help
```

**Authenticate (GitHub device-code flow):**

```bash
mcp-publisher login github
# Follow the prompt: open https://github.com/login/device and enter the displayed code.
# This grants access to the io.github.salahawad/* namespace.
```

**Publish:**

```bash
cd /path/to/outlook-personal-mcp
mcp-publisher publish
# Reads server.json and pushes metadata to the registry.
```

The `name` in `server.json` (`io.github.salahawad/outlook-personal-mcp`) must match the `<!-- mcp-name: ... -->` marker in the README (which PyPI surfaces as the package's long-description). The version in `server.json` and `packages[0].version` must match the published PyPI release.

---

## 4. Build and Submit the MCPB Desktop Extension

MCPB (MCP Bundle) packages an MCP server as a self-contained desktop extension (`.mcpb` file) for Claude Desktop.

**Install the `mcpb` CLI:**

```bash
npm install -g @anthropic-ai/mcpb
mcpb --help
```

**Build the bundle:**

```bash
cd /path/to/outlook-personal-mcp
mcpb pack
# Produces outlook-personal-mcp-<version>.mcpb in the current directory.
```

`mcpb pack` reads `manifest.json` from the repo root. The `server.type = "uv"` entry tells the Desktop host to manage Python and the package's dependencies via `uv` — no bundling of `site-packages` is required.

**Submission checklist before packing:**

| Requirement | Status |
|---|---|
| `manifest.json` present at repo root | Done |
| `manifest_version` is `"0.4"` | Done |
| Tool annotations present (readOnlyHint, destructiveHint, etc.) | Done — all 31 tools annotated |
| Privacy policy at a publicly reachable URL | Done — `PRIVACY.md` on `main` branch |
| `privacy_policies` array in `manifest.json` points to it | Done |
| Extension tested on Windows and macOS | **TODO** — test before submitting |
| Reviewer test account info prepared | See below |

**Reviewer test account:**

The reviewer needs to install Claude Desktop, run the server, and exercise the tools. To let them do this:

1. Create a sandbox Microsoft personal account (Outlook.com) with some test emails and calendar events.
2. Register a free Azure app for the sandbox account following the README's "Azure App Registration" steps.
3. Include the client ID and a note that the reviewer will need to run `mcp-outlook-personal login` once to complete the device-code flow.

**Submit:**

Submit the `.mcpb` file via the Google Form: [https://clau.de/desktop-extention-submission](https://clau.de/desktop-extention-submission)

Include a short description, the reviewer test account instructions above, and a link to the public `PRIVACY.md`.

**Dependency note:**

This extension uses `server.type = "uv"` in `manifest.json`. This means Claude Desktop will invoke `uv run --with mcp-outlook-personal mcp-outlook-personal` to launch the server, and `uv` handles downloading and caching the package and all its dependencies (`mcp`, `msal`, `httpx`, `pydantic`). No manual vendoring into a `lib/` directory is required. The host machine must have `uv` installed (or Claude Desktop must ship it); if the Desktop host does not bundle `uv`, the extension will fail to launch — this is the main portability concern for the `uv` runtime type.
