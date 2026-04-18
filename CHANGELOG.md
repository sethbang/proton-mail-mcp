# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.0] - 2026-04-17

### Added

- **`get_thread` tool** — find all messages in a conversation by walking In-Reply-To and References headers. Returns messages sorted chronologically (oldest first).
- **`mark_all_read` tool** — bulk mark all unread messages in a folder as read, with optional `olderThan` date filter. Single IMAP connection + single STORE command.
- **`save_draft` tool** — save an email as a draft without sending it. Uses IMAP APPEND to place a fully-formed message in the Drafts folder for user review.
- `send_email` now attempts a best-effort lookup of the sent copy UID in the Sent folder and includes it in the response when found.

### Changed

- **HTML-to-text conversion** now uses the `html-to-text` package instead of naive regex stripping. Preserves links as `text [url]`, formats lists with bullets, skips tracking pixels, and handles entities correctly.

### Fixed

- **Search and list ordering is now correct by date.** Replaced the 5x UID heuristic with a shared `fetchSortAndLimit()` helper that fetches all matching envelopes (up to 500) before sorting by date. This fixes the reported bug where `search_messages` and `list_messages` could return old messages instead of the most recent ones when UIDs didn't correlate with dates (e.g. moved messages).

## [0.3.0] - 2025-04-15

### Added

- `send_email`, `reply_email`, and `forward_email` now return the Message-ID in their response, allowing callers to correlate sent messages with what appears in the mailbox.
- `reply_email` includes quoted original message with attribution line (`On <date>, <sender> wrote:`). Opt out with `includeQuote: false`.
- `move_message` and `delete_message` (non-permanent) now return the new UID in the destination folder when the server supports the UIDPLUS extension.
- Truncation marker includes the original body length and a suggested `maxBodyLength` value to retrieve the full message.

### Fixed

- `list_messages` ordering is now consistent across back-to-back calls. Replaced racy sequence-number-based fetch with UID-based search, eliminating cases where recent messages could be missing or returned in inconsistent order.

## [0.2.1] - 2025-03-22

### Fixed

- Handle single-part messages where the IMAP body structure has no `part` field (defaults to part "1").
- Add PGP fallback: when no `text/plain` or `text/html` part is found (e.g. PGP-encrypted messages), download part "1" as raw text instead of returning an empty body.

## [0.2.0] - 2025-03-22

### Added

- npm trusted publisher workflow for provenance-signed releases.

### Fixed

- Decode quoted-printable and base64 transfer encodings via `client.download()` instead of raw `source` fetch.
- Prefer `text/plain` part over `text/html` when both are available in multipart messages.
- Truncate large message bodies to 50,000 characters by default (configurable via `maxBodyLength`).
- Sort `list_messages` results by date descending with UID tiebreak.

## [0.1.0] - 2025-03-17

### Added

- Initial release.
- **SMTP tools:** `send_email` with attachments, `reply_email` with threading headers, `forward_email`.
- **IMAP tools:** `list_folders`, `list_messages` with UID-based pagination, `read_message` with HTML stripping, `search_messages`, `download_attachment`.
- **Mailbox management:** `move_message`, `delete_message` (soft-delete to Trash by default), `update_message_flags`.
- Safety protections: soft-delete, MCP tool annotations, `READONLY` mode.
- Security hardening: input validation, credential sanitization in error messages, rate limiting (10 emails/min), attachment size limits.
- Debug logging to stderr via `DEBUG=true`.

[0.4.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.2.1...v0.3.0
[0.2.1]: https://github.com/sethbang/proton-mail-mcp/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/sethbang/proton-mail-mcp/releases/tag/v0.1.0
