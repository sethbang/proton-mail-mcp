# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.5.0] - 2026-04-20

### Added

- **`list_attachments` tool** — fetch attachment metadata (part numbers, filenames, types, sizes) without downloading the body. Composes with `download_attachment` for bulk extraction workflows.
- **`get_thread.messageId` input** — preferred over UID+folder; walks INBOX, Sent, and All Mail by default so cross-folder threads (the common case) are returned intact. New optional `folders` parameter lets callers override the default set. Output rows are tagged `UID X @FolderName` to disambiguate per-folder copies.
- **`read_message.showHeaders`** — when `true`, surfaces `In-Reply-To`, `References`, `Reply-To`, `List-Unsubscribe`, and `List-ID` under an `--- Extra Headers ---` section. Handles RFC 5322 folded continuation lines.
- **`read_message.stripUrls`** — when `true`, drops anchor URLs from stripped-HTML output so agents summarizing newsletters don't burn tokens on tracking URLs.
- **`save_draft.fromName` and `save_draft.replyTo`** — feature parity with `send_email`.
- **`forward_email.includeAttachments`** — defaults to `true` so forwards now carry original attachments as agents expect; set to `false` to forward the body alone.
- **`list_messages` pagination hint** — when the result count equals the requested limit, the response footer suggests the `beforeUid` value to use for the next page.

### Fixed

- **Silent false success on bogus UIDs** for `delete_message`, `update_message_flags`, and `move_message`. IMAP STORE/DELETE/MOVE succeed silently against missing UIDs; we now pre-check existence via `fetchOne` and throw `"Message UID N not found in <folder>"` (matching the pattern `read_message` already used). Applies equally to `list_attachments` and `download_attachment`.
- **Attachment content returned as body** in `read_message` for HTML messages that carried a `text/plain` attachment. The body-part selector previously matched on content-type only; it now skips parts with `Content-Disposition: attachment` so the body is the real body.
- **`forward_email` dropped original attachments.** Forwards now download each attachment from the source message and re-attach.
- **Folder-not-found errors are actionable across every folder-taking tool.** All `getMailboxLock` calls route through a shared helper that translates imapflow's canonical `err.mailboxMissing` / `err.serverResponseCode ∈ {NONEXISTENT, TRYCREATE}` signals into `"Folder not found: <path>"`. `move_message` destination failures now probe the destination with `client.status()` to distinguish missing folders from other failures. `saveDraft` append failures are translated the same way.
- **Flag whitelist in `validateImapFlag`.** Previously `\`-prefixed names were accepted wholesale, so IMAP would silently drop unknown system flags like `\Bogus`. Now we accept only RFC 3501 system flags (`\Seen`, `\Flagged`, `\Answered`, `\Draft`, `\Deleted`, `\Recent`) plus alphanumeric user keywords per RFC 3501. The error message lists the allowed set.
- **Date string validation** for `search_messages.since` / `search_messages.before` / `mark_all_read.olderThan`. New `isValidDateString` helper rejects anything that isn't a valid `YYYY-MM-DD` calendar date (catches `04-19-2026`, `2026-13-01`, `2026-02-29`, etc.).
- **`download_attachment` errors are now actionable.** A bogus part number throws `"Part X not found on UID Y in <folder>; known parts: [2, 3]"` (or `[(none)]` when the message has no attachments) instead of imapflow's opaque `"Command failed"`.
- **`mark_all_read` copy under `olderThan`.** When the filter matches zero messages, the response now says `"No unread messages older than <date> in <folder>"` instead of implying the folder has no unread mail.
- **BCC addresses are no longer echoed verbatim** in the `send_email` success response. Masked to `"and BCC to N recipient(s)"` so verbatim-logged output doesn't leak the BCC list.

### Changed

- Tool descriptions updated to document semantics that were previously implicit:
  - `update_message_flags` now lists both RFC 3501 system flags and user keywords.
  - `send_email.replyTo` and `save_draft.replyTo` note that Proton SMTP may rewrite non-authenticated values.
  - `search_messages.since` is documented as inclusive; `search_messages.before` and `mark_all_read.olderThan` as exclusive.
  - `read_message.maxBodyLength` surfaces its 100..500 000 range in prose.

## [0.4.1] - 2026-04-19

### Fixed

- **Critical: `list_messages` and `search_messages` no longer fail with "Connection not available"** on real IMAP connections. The v0.4.0 refactor introduced `return this.fetchSortAndLimit(...)` without `await`, which caused the outer `try/finally` to close the IMAP connection (via `client.logout()`) before the fetch iterator resolved. Unit tests passed because mocks have no real connection state. Fixed by adding `await` to both call sites and adding regression tests that simulate connection closure mid-fetch.

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

[0.5.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.4.1...v0.5.0
[0.4.1]: https://github.com/sethbang/proton-mail-mcp/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.2.1...v0.3.0
[0.2.1]: https://github.com/sethbang/proton-mail-mcp/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/sethbang/proton-mail-mcp/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/sethbang/proton-mail-mcp/releases/tag/v0.1.0
