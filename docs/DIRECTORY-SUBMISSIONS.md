# Directory submission checklist

Where to list the server for discoverability, and how each one works (verified
2026). Do these **after** the official MCP Registry publish — several directories
ingest the registry automatically, so that one step seeds the rest.

## 0. Official MCP Registry (do first — seeds the others)

`modelcontextprotocol/servers` retired its third-party list; the
[official registry](https://registry.modelcontextprotocol.io) is now the path.
The manifest (`server.json`) and the `mcpName` field in `package.json` are
already in the repo (added in the registry-prep PR). Remaining manual steps —
they need interactive auth, so run them yourself:

```bash
npm publish --access public        # republish so the npm package carries mcpName
mcp-publisher login github         # device-code OAuth; namespace = io.github.sethbang/*
mcp-publisher publish              # reads ./server.json
# verify:
curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.sethbang/proton-mail-mcp"
```

Publishing here **auto-propagates** to **PulseMCP** (daily ingest) and the
**Glama** registry superset — no separate submission needed for those.

## 1. Smithery — https://smithery.ai

- **Method:** connect the GitHub repo via the Smithery GitHub App, or use the
  CLI. Needs a `smithery.yaml` at repo root describing a stdio start command.
- **Submit:** https://smithery.ai/new
- **Checklist:** add `smithery.yaml` → install the GitHub App / connect repo →
  deploy.

## 2. Glama — https://glama.ai/mcp

- **Method:** auto-indexed from GitHub (and from the official registry), plus a
  manual "Add Server" button. You can **claim** the listing as the author for an
  admin panel. A clean README ranks better.
- **Submit:** https://glama.ai/mcp/servers → "Add Server" (paste the repo URL) →
  then claim ownership.

## 3. mcp.so — https://mcp.so

- **Method:** manual — "Submit" in the nav, or open an issue on `chatmcp/mcpso`.
- **Submit:** https://mcp.so/submit
- **Provide:** name, one-line description, tool count (~30), transport (stdio),
  GitHub URL, homepage (the Pages site), optional icon, and a client config
  snippet.

## 4. awesome-mcp-servers (punkpeye) — feeds mcpservers.org

- **Method:** PR to https://github.com/punkpeye/awesome-mcp-servers (metadata
  only: repo link, description, language, platform). Follow its `CONTRIBUTING.md`
  for the exact entry format and ordering.
- There's also a form at https://mcpservers.org/submit.

## Suggested entry copy (reuse across forms)

> **Proton Mail (unofficial)** — Send, read, search, and organize Proton Mail
> from any MCP client over SMTP + IMAP (via Proton Mail Bridge). Privacy-focused
> and safety-hardened: read-only mode, dry-run previews on bulk ops,
> Trash-by-default deletes, outbound HTML sanitization. Node/TypeScript, MIT.
> Not affiliated with Proton AG.

## Repo metadata (one-time, in the GitHub UI / gh CLI)

- [ ] **Homepage** → the GitHub Pages site (set via `gh repo edit --homepage`).
- [ ] **About description** → prefixed with "Unofficial." (done).
- [ ] **Topics** → already good (`mcp`, `mcp-server`, `proton-mail`, …).
- [ ] **Social preview image** → crop a frame from the demo GIF to 1280×640
      (Settings → General → Social preview).
- [ ] Add a couple of `good first issue` / `help wanted` labeled issues so
      drive-by contributors have an entry point.
