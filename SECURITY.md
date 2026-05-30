# Security Policy

Proton Mail MCP is an independent, community project (unofficial — not affiliated
with Proton AG). It handles email credentials and can read, send, and delete
mail, so security reports are taken seriously.

> For vulnerabilities in **Proton Mail**, **Proton Mail Bridge**, or any Proton
> product itself, report to Proton — not here. See
> <https://proton.me/security/disclosure>. This policy covers only this MCP
> server.

## Reporting a vulnerability

**Please do not open a public GitHub issue for security vulnerabilities.**

Use **GitHub's private vulnerability reporting**:

1. Go to the repository's **Security** tab →
   **[Report a vulnerability](https://github.com/sethbang/proton-mail-mcp/security/advisories/new)**.
2. Describe the issue, affected version, and steps to reproduce.

This routes the report privately to the maintainer. You'll get an
acknowledgement, and once a fix ships, credit in the advisory if you'd like it.

## Scope

In scope and useful to report:

- Credential leakage (e.g. secrets written to logs, error messages, or tool
  responses).
- Path-traversal or arbitrary file read/write via the attachment download tools.
- Prompt-injection paths where untrusted email content can drive the server into
  an unintended action (beyond the known, documented surface).
- Bypasses of the safety controls (`READONLY`, `RESTRICT_OUTBOUND_TO_SELF`, the
  `fromName` spoofing guard, HTML sanitization).

## Out of scope

- Vulnerabilities in Proton Mail / Proton Mail Bridge themselves (report to
  Proton, link above).
- Issues that require the operator to deliberately misconfigure the server in a
  way the docs warn against (e.g. enabling `ALLOW_EMPTY_FOLDER`).
- Vulnerabilities in third-party dependencies already tracked upstream
  (Dependabot handles routine dependency bumps).

## Handling your own credentials

A reminder for operators, not a vulnerability class: use a Proton **SMTP
password** (the app-specific one), not your main login password; keep it in your
MCP client's env config, not in source control.
