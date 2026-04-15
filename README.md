# Proton Mail MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that gives AI assistants full access to your Proton Mail account -- send, read, search, and organize email over SMTP and IMAP.

## Features

- **Send email** via Proton Mail SMTP (direct or through Proton Mail Bridge)
- **Read email** via IMAP through [Proton Mail Bridge](https://proton.me/mail/bridge)
- **Search** messages by sender, recipient, subject, body, date, and flags
- **List folders** with message and unread counts
- Works with any MCP-compatible client (Claude Desktop, Claude Code, Cursor, etc.)

## Prerequisites

- [Node.js](https://nodejs.org/) v18+
- A [Proton Mail](https://proton.me/mail) account with an SMTP password ([how to get one](https://proton.me/support/smtp-submission))
- [Proton Mail Bridge](https://proton.me/mail/bridge) running locally (required for IMAP/read tools)

## Quick Start

```bash
git clone https://github.com/sethbang/proton-mail-mcp.git
cd proton-mail-mcp
npm install
npm run build
```

Copy `.env.example` to `.env` and fill in your credentials (see [Configuration](#configuration) below).

### Add to your MCP client

Add the following to your client's MCP server configuration:

```json
{
  "mcpServers": {
    "protonmail": {
      "command": "node",
      "args": ["/absolute/path/to/proton-mail-mcp/build/index.js"],
      "env": {
        "PROTONMAIL_USERNAME": "your-email@protonmail.com",
        "PROTONMAIL_PASSWORD": "your-smtp-password"
      }
    }
  }
}
```

## Tools

### `send_email`

Send an email using Proton Mail SMTP.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `to` | Yes | Recipient address(es), comma-separated |
| `subject` | Yes | Subject line |
| `body` | Yes | Plain text or HTML content |
| `isHtml` | No | Whether `body` is HTML (default: `false`) |
| `cc` | No | CC recipient(s), comma-separated |
| `bcc` | No | BCC recipient(s), comma-separated |
| `replyTo` | No | Reply-To address |
| `fromName` | No | Display name for the From field |

### `list_folders`

List all mailbox folders with message and unread counts. No parameters.

### `list_messages`

List recent messages from a folder.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `folder` | No | Folder path (default: `INBOX`) |
| `limit` | No | Max messages to return, 1-100 (default: `20`) |

### `read_message`

Read a specific message by UID. Returns full headers and body.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `uid` | Yes | Message UID (from `list_messages` or `search_messages`) |
| `folder` | No | Folder path (default: `INBOX`) |

### `search_messages`

Search messages by various criteria.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `folder` | No | Folder to search (default: `INBOX`) |
| `from` | No | Filter by sender |
| `to` | No | Filter by recipient |
| `subject` | No | Filter by subject (substring match) |
| `body` | No | Filter by body content (substring match) |
| `since` | No | Messages since date (`YYYY-MM-DD`) |
| `before` | No | Messages before date (`YYYY-MM-DD`) |
| `seen` | No | `true` = read, `false` = unread |
| `flagged` | No | Filter by flagged/starred status |
| `limit` | No | Max results, 1-100 (default: `20`) |

## Configuration

### Environment Variables

**SMTP (required):**

| Variable | Default | Description |
|----------|---------|-------------|
| `PROTONMAIL_USERNAME` | -- | Your Proton Mail email address |
| `PROTONMAIL_PASSWORD` | -- | Your SMTP password ([not your login password](https://proton.me/support/smtp-submission)) |
| `PROTONMAIL_HOST` | `smtp.protonmail.ch` | SMTP host |
| `PROTONMAIL_PORT` | `587` | SMTP port |
| `PROTONMAIL_SECURE` | `false` | Use TLS (`true` for port 465) |

**IMAP (for read/search tools):**

| Variable | Default | Description |
|----------|---------|-------------|
| `IMAP_HOST` | `127.0.0.1` | Proton Mail Bridge host |
| `IMAP_PORT` | `1143` | Bridge IMAP port |
| `IMAP_SECURE` | `false` | Use TLS |
| `IMAP_USERNAME` | _falls back to SMTP username_ | Bridge username |
| `IMAP_PASSWORD` | _falls back to SMTP password_ | Bridge password |

**Debug:**

| Variable | Default | Description |
|----------|---------|-------------|
| `DEBUG` | `false` | Enable verbose logging to stderr |

## Development

```bash
npm run build          # Compile TypeScript
npm run watch          # Compile in watch mode
npm run test           # Run tests
npm run test:watch     # Run tests in watch mode
npm run lint           # Lint with ESLint
npm run format         # Format with Prettier
npm run inspector      # Launch MCP inspector
```

## Acknowledgments

Originally based on [protonmail-mcp](https://github.com/amotivv/protonmail-mcp) by [amotivv, inc.](https://amotivv.com)

## License

MIT -- see [LICENSE](LICENSE) for details.
