/**
 * Shared validation and sanitization utilities for the MCP server.
 * Centralizes security-sensitive input handling so it can be tested in isolation.
 */

/**
 * Validate an IMAP folder path. Rejects control characters, CRLF sequences,
 * and excessively long names. Returns the trimmed folder path or throws.
 */
export function validateFolderPath(folder: string): string {
  const trimmed = folder.trim();
  if (trimmed.length === 0) {
    throw new Error("Folder path must not be empty");
  }
  if (trimmed.length > 500) {
    throw new Error("Folder path exceeds maximum length of 500 characters");
  }
  // Reject control characters (0x00-0x1f, 0x7f)
  if (/[\x00-\x1f\x7f]/.test(trimmed)) {
    throw new Error("Folder path contains invalid control characters");
  }
  return trimmed;
}

/**
 * Validate a MIME part number (e.g. "1", "1.2", "1.2.3").
 * Returns the part number or throws on invalid format.
 */
export function validatePartNumber(partNumber: string): string {
  if (!/^\d+(\.\d+)*$/.test(partNumber)) {
    throw new Error(`Invalid MIME part number: ${partNumber}`);
  }
  return partNumber;
}

/**
 * Sanitize an attachment filename by stripping path separators and control characters.
 * Returns a safe basename, falling back to "attachment" if nothing remains.
 */
export function sanitizeFilename(filename: string): string {
  // Take only the last path segment (basename)
  const basename = filename.replace(/^.*[/\\]/, "");
  // Strip control characters and null bytes
  const cleaned = basename.replace(/[\x00-\x1f\x7f]/g, "");
  return cleaned || "attachment";
}

/**
 * Sanitize a display name for the From header. Strips characters that could
 * enable header injection or break RFC 5322 quoted-string format.
 */
export function sanitizeFromName(name: string): string {
  return name.replace(/["\\<>\r\n]/g, "");
}

/**
 * Validate a single email address. Stricter than a basic format check:
 * rejects characters that could enable header injection.
 */
export function isValidEmailAddress(addr: string): boolean {
  // Reject dangerous characters outright
  if (/[<>";\r\n]/.test(addr)) {
    return false;
  }
  // RFC 5321 local-part + domain check
  return /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/.test(
    addr,
  );
}

/**
 * Validate an IMAP flag string. Accepts system flags (e.g. \Seen, \Flagged)
 * and custom keywords (alphanumeric + underscore). Throws on invalid.
 */
export function validateImapFlag(flag: string): string {
  // System flags: backslash followed by letters
  if (/^\\[A-Za-z]+$/.test(flag)) {
    return flag;
  }
  // Custom keywords: alphanumeric + underscore
  if (/^[A-Za-z0-9_]+$/.test(flag)) {
    return flag;
  }
  throw new Error(`Invalid IMAP flag: ${flag}`);
}

/**
 * Produce a safe error message for MCP clients. Returns the error class name
 * and a generic category instead of the raw message, which may contain credentials.
 * Full error details are logged to stderr separately.
 */
export function sanitizeErrorMessage(error: unknown): string {
  if (!(error instanceof Error)) {
    return "Unknown error";
  }

  const name = error.constructor.name || "Error";

  // Categorize by common error patterns
  const msg = error.message.toLowerCase();
  if (msg.includes("timeout") || msg.includes("timed out")) {
    return `${name}: connection timed out`;
  }
  if (msg.includes("auth") || msg.includes("login") || msg.includes("credential")) {
    return `${name}: authentication failed`;
  }
  if (msg.includes("econnrefused") || msg.includes("enotfound") || msg.includes("connect")) {
    return `${name}: connection failed`;
  }
  if (msg.includes("not found")) {
    return `${name}: ${error.message}`;
  }

  // For non-sensitive errors (validation, etc.), pass through a truncated message
  // but strip anything that looks like a credential
  const sanitized = error.message
    .replace(/(?:user|pass|password|auth|credentials?)[=:\s]+"?[^\s"]+/gi, "[REDACTED]")
    .replace(/(?:smtp|imap):\/\/[^\s]+/gi, "[REDACTED_URL]")
    .slice(0, 200);
  return `${name}: ${sanitized}`;
}
