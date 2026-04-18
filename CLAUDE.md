# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MCP (Model Context Protocol) server for Proton Mail email. Exposes tools for sending emails via SMTP and reading emails via IMAP (through Proton Mail Bridge). Communicates over stdio transport.

## Commands

- `npm run build` — compile TypeScript to `build/`
- `npm run watch` — compile in watch mode
- `npm run inspector` — launch the MCP inspector against the built server
- `npm run lint` — run ESLint
- `npm run test` — run Vitest unit tests
- `npm run test:watch` — run tests in watch mode
- `npm run format` — format with Prettier
- `npm run format:check` — check formatting

## Architecture

Four source files:

- `src/index.ts` — MCP server entry point. Reads SMTP and IMAP config from environment variables, registers all tool handlers via `McpServer.registerTool()` with Zod input schemas, starts stdio transport. SMTP verification on startup is non-fatal (warns and continues).
- `src/email-service.ts` — `EmailService` class wrapping nodemailer. Handles transporter creation with timeouts, `sendEmail()` (returns `SMTPTransport.SentMessageInfo` with `messageId`), `buildRawMessage()` (for draft APPEND), and `verifyConnection()`.
- `src/imap-service.ts` — `ImapService` class wrapping imapflow. Provides `listFolders()`, `listMessages()`, `readMessage()`, `searchMessages()`, `getThread()`, `moveMessage()` (returns `MoveResult` with new UID), `deleteMessage()`, `updateFlags()`, `markAllRead()`, `findByMessageId()`, `saveDraft()`, and `downloadAttachment()`. Uses `fetchSortAndLimit()` to guarantee date-based ordering. Each operation creates a fresh IMAP connection to avoid stale connection issues.
- `src/validation.ts` — Input validation and sanitization utilities: folder path, MIME part number, email address, IMAP flag validation; filename and from-name sanitization; error message sanitization for MCP clients.

Tests live in `src/__tests__/` and are excluded from the TypeScript build via `tsconfig.json`.

## MCP Tools

**SMTP (sending) — all return Message-ID in response:**
- `send_email` — send email, returns Message-ID + best-effort Sent folder UID
- `reply_email` — reply with threading headers, quoted original (opt out via `includeQuote: false`)
- `forward_email` — forward with original content and threading headers
- `save_draft` — save email as draft via IMAP APPEND (does not send)

**IMAP (reading, requires Proton Mail Bridge):**
- `list_folders` — list mailbox folders with message/unread counts
- `list_messages` — list recent messages sorted by date (newest first), UID-based pagination
- `read_message` — read a specific message by UID, returns headers and body (truncation shows original length)
- `search_messages` — search by from/to/subject/body/date/flags, sorted by date (newest first)
- `get_thread` — find all messages in a conversation thread via References/In-Reply-To headers
- `download_attachment` — download attachment by MIME part number (base64)

**Mailbox management — move/delete return new UID when UIDPLUS is supported:**
- `move_message` — move message to a different folder
- `delete_message` — soft-delete to Trash (default) or permanent expunge
- `update_message_flags` — add/remove flags (\\Seen, \\Flagged, etc.)
- `mark_all_read` — bulk mark all unread messages in a folder as read

## Required Environment Variables

**SMTP:** `PROTONMAIL_USERNAME`, `PROTONMAIL_PASSWORD` (SMTP password, not login password). Optional: `PROTONMAIL_HOST` (default `smtp.protonmail.ch`), `PROTONMAIL_PORT` (default `587`), `PROTONMAIL_SECURE` (default `false`).

**IMAP:** `IMAP_HOST` (default `127.0.0.1`), `IMAP_PORT` (default `1143`), `IMAP_SECURE` (default `false`). `IMAP_USERNAME` and `IMAP_PASSWORD` default to the SMTP credentials if not set.

`DEBUG` — enable verbose stderr logging.

`READONLY` — when `true`, disables all mutating tools (send, reply, forward, move, delete, flags). Only read/search tools are registered.

See `.env.example` for a template.

## Notes

- Uses MCP SDK v1.29.0 — tools are registered via `McpServer.registerTool()` with Zod schemas for input validation. Input is parsed/typed by Zod before reaching the handler.
- ESM project (`"type": "module"`) with Node16 module resolution. Imports must use `.js` extensions.
- Debug/error logging goes to stderr (required for MCP stdio servers — stdout is the protocol channel).
- Error messages returned to MCP clients are sanitized via `sanitizeError()` to strip credential-like substrings. Full errors are logged to stderr only.
- IMAP tools require Proton Mail Bridge running locally. Bridge decrypts mail and exposes a local IMAP server.
