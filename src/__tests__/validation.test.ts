import { describe, it, expect } from "vitest";
import * as path from "node:path";
import {
  validateFolderPath,
  validatePartNumber,
  sanitizeFilename,
  sanitizeFromName,
  isValidEmailAddress,
  validateImapFlag,
  sanitizeErrorMessage,
  isValidDateString,
  validateBulkInput,
  permanentDeleteNeedsConfirm,
  extractEmailAddress,
  buildReplyAllRecipients,
  externalRecipientAddresses,
  validateSizeBound,
  validateListId,
  resolveAllowlistedPath,
  validateSubject,
  sanitizeEmailHtml,
  detectActiveHtml,
  snippetNote,
  SNIPPET_UNTRUSTED_NOTE,
  isValidBase64,
  isValidMimeContentType,
} from "../validation.js";

describe("validateFolderPath", () => {
  it("accepts standard folder names", () => {
    expect(validateFolderPath("INBOX")).toBe("INBOX");
    expect(validateFolderPath("Sent")).toBe("Sent");
    expect(validateFolderPath("INBOX/subfolder")).toBe("INBOX/subfolder");
    expect(validateFolderPath("INBOX.Drafts")).toBe("INBOX.Drafts");
    expect(validateFolderPath("[Gmail]/All Mail")).toBe("[Gmail]/All Mail");
  });

  it("trims whitespace", () => {
    expect(validateFolderPath("  INBOX  ")).toBe("INBOX");
  });

  it("rejects empty string", () => {
    expect(() => validateFolderPath("")).toThrow("must not be empty");
    expect(() => validateFolderPath("   ")).toThrow("must not be empty");
  });

  it("rejects strings exceeding max length", () => {
    const long = "A".repeat(501);
    expect(() => validateFolderPath(long)).toThrow("exceeds maximum length");
  });

  it("rejects control characters", () => {
    expect(() => validateFolderPath("INBOX\x00")).toThrow("invalid control characters");
    expect(() => validateFolderPath("INBOX\r\nDELETE")).toThrow("invalid control characters");
    expect(() => validateFolderPath("IN\tBOX")).toThrow("invalid control characters");
    expect(() => validateFolderPath("INBOX\x7f")).toThrow("invalid control characters");
  });

  it("accepts folder names with unicode characters", () => {
    expect(validateFolderPath("Brouillons")).toBe("Brouillons");
    expect(validateFolderPath("Gesendet")).toBe("Gesendet");
  });
});

describe("validatePartNumber", () => {
  it("accepts valid MIME part numbers", () => {
    expect(validatePartNumber("1")).toBe("1");
    expect(validatePartNumber("1.2")).toBe("1.2");
    expect(validatePartNumber("1.2.3")).toBe("1.2.3");
    expect(validatePartNumber("10.20.30")).toBe("10.20.30");
  });

  it("rejects invalid part numbers", () => {
    expect(() => validatePartNumber("")).toThrow("Invalid MIME part number");
    expect(() => validatePartNumber("abc")).toThrow("Invalid MIME part number");
    expect(() => validatePartNumber("1.2.3a")).toThrow("Invalid MIME part number");
    expect(() => validatePartNumber("../1")).toThrow("Invalid MIME part number");
    expect(() => validatePartNumber("../../etc/passwd")).toThrow("Invalid MIME part number");
    expect(() => validatePartNumber("1; DROP TABLE")).toThrow("Invalid MIME part number");
  });
});

describe("sanitizeFilename", () => {
  it("passes through clean filenames", () => {
    expect(sanitizeFilename("report.pdf")).toBe("report.pdf");
    expect(sanitizeFilename("my file (1).docx")).toBe("my file (1).docx");
  });

  it("strips path traversal sequences", () => {
    expect(sanitizeFilename("../../etc/passwd")).toBe("passwd");
    expect(sanitizeFilename("..\\..\\windows\\system32\\config")).toBe("config");
    expect(sanitizeFilename("/absolute/path/file.txt")).toBe("file.txt");
  });

  it("strips control characters", () => {
    expect(sanitizeFilename("file\x00name.txt")).toBe("filename.txt");
    expect(sanitizeFilename("file\x1fname.txt")).toBe("filename.txt");
    expect(sanitizeFilename("file\x7fname.txt")).toBe("filename.txt");
  });

  it("falls back to 'attachment' when nothing remains", () => {
    expect(sanitizeFilename("/")).toBe("attachment");
    expect(sanitizeFilename("\\")).toBe("attachment");
    expect(sanitizeFilename("\x00\x01")).toBe("attachment");
  });
});

describe("sanitizeFromName", () => {
  it("passes through clean names", () => {
    expect(sanitizeFromName("John Doe")).toBe("John Doe");
    expect(sanitizeFromName("O'Brien")).toBe("O'Brien");
  });

  it("strips quotes and backslashes", () => {
    expect(sanitizeFromName('John "Bobby" Doe')).toBe("John Bobby Doe");
    expect(sanitizeFromName("back\\slash")).toBe("backslash");
  });

  it("strips angle brackets", () => {
    expect(sanitizeFromName("Evil <script>")).toBe("Evil script");
  });

  it("strips CRLF to prevent header injection", () => {
    expect(sanitizeFromName("Evil\r\nBcc: attacker@evil.com")).toBe("EvilBcc: attacker@evil.com");
    expect(sanitizeFromName("Name\nHeader: value")).toBe("NameHeader: value");
  });
});

describe("isValidEmailAddress", () => {
  it("accepts valid addresses", () => {
    expect(isValidEmailAddress("user@example.com")).toBe(true);
    expect(isValidEmailAddress("user+tag@example.com")).toBe(true);
    expect(isValidEmailAddress("user.name@sub.domain.com")).toBe(true);
    expect(isValidEmailAddress("user@protonmail.com")).toBe(true);
  });

  it("rejects addresses with dangerous characters", () => {
    expect(isValidEmailAddress('"><script>@x.y')).toBe(false);
    expect(isValidEmailAddress("user@example.com;drop")).toBe(false);
    expect(isValidEmailAddress("user<@example.com")).toBe(false);
    expect(isValidEmailAddress("user>@example.com")).toBe(false);
  });

  it("rejects addresses with CRLF", () => {
    expect(isValidEmailAddress("user\r\n@example.com")).toBe(false);
    expect(isValidEmailAddress("user@example\r\n.com")).toBe(false);
  });

  it("rejects malformed addresses", () => {
    expect(isValidEmailAddress("")).toBe(false);
    expect(isValidEmailAddress("noatsign")).toBe(false);
    expect(isValidEmailAddress("@nodomain")).toBe(false);
    expect(isValidEmailAddress("user@")).toBe(false);
    expect(isValidEmailAddress("user@.com")).toBe(false);
  });
});

describe("validateImapFlag", () => {
  it("accepts system flags", () => {
    expect(validateImapFlag("\\Seen")).toBe("\\Seen");
    expect(validateImapFlag("\\Flagged")).toBe("\\Flagged");
    expect(validateImapFlag("\\Answered")).toBe("\\Answered");
    expect(validateImapFlag("\\Draft")).toBe("\\Draft");
    expect(validateImapFlag("\\Deleted")).toBe("\\Deleted");
  });

  it("accepts custom keyword flags", () => {
    expect(validateImapFlag("Important")).toBe("Important");
    expect(validateImapFlag("Custom_Tag")).toBe("Custom_Tag");
    expect(validateImapFlag("label123")).toBe("label123");
  });

  it("rejects invalid flags", () => {
    expect(() => validateImapFlag("")).toThrow("Invalid IMAP flag");
    expect(() => validateImapFlag("has spaces")).toThrow("Invalid IMAP flag");
    expect(() => validateImapFlag("\\")).toThrow("Invalid IMAP flag");
    expect(() => validateImapFlag("flag\r\n")).toThrow("Invalid IMAP flag");
    expect(() => validateImapFlag("flag;injection")).toThrow("Invalid IMAP flag");
  });

  it("rejects backslash-prefixed flags that aren't RFC 3501 system flags", () => {
    // IMAP servers silently drop unknown system flags. Catch these at the client boundary
    // rather than letting the agent believe "flag added" when nothing happened.
    expect(() => validateImapFlag("\\Bogus")).toThrow("Invalid IMAP flag");
    expect(() => validateImapFlag("\\NotReal")).toThrow("Invalid IMAP flag");
    expect(() => validateImapFlag("\\Xyz")).toThrow("Invalid IMAP flag");
  });

  it("error message for unknown system flag lists the allowed set", () => {
    try {
      validateImapFlag("\\Bogus");
      throw new Error("should have thrown");
    } catch (err) {
      const msg = (err as Error).message;
      expect(msg).toContain("\\Seen");
      expect(msg).toContain("\\Flagged");
      expect(msg).toContain("\\Answered");
      expect(msg).toContain("\\Draft");
      expect(msg).toContain("\\Deleted");
    }
  });

  it("accepts \\Recent as a system flag", () => {
    expect(validateImapFlag("\\Recent")).toBe("\\Recent");
  });
});

describe("isValidDateString", () => {
  it("accepts valid YYYY-MM-DD dates", () => {
    expect(isValidDateString("2026-04-19")).toBe(true);
    expect(isValidDateString("2020-01-01")).toBe(true);
    expect(isValidDateString("1999-12-31")).toBe(true);
    expect(isValidDateString("2000-02-29")).toBe(true); // leap year
  });

  it("rejects wrong format", () => {
    expect(isValidDateString("04-19-2026")).toBe(false);
    expect(isValidDateString("2026/04/19")).toBe(false);
    expect(isValidDateString("2026.04.19")).toBe(false);
    expect(isValidDateString("19-04-2026")).toBe(false);
  });

  it("rejects invalid calendar dates", () => {
    expect(isValidDateString("2026-13-01")).toBe(false); // month 13
    expect(isValidDateString("2026-00-01")).toBe(false); // month 0
    expect(isValidDateString("2026-04-32")).toBe(false); // day 32
    expect(isValidDateString("2026-04-00")).toBe(false); // day 0
    expect(isValidDateString("2025-02-29")).toBe(false); // not a leap year
  });

  it("rejects empty and non-date strings", () => {
    expect(isValidDateString("")).toBe(false);
    expect(isValidDateString("today")).toBe(false);
    expect(isValidDateString("yesterday")).toBe(false);
    expect(isValidDateString("2026-04")).toBe(false);
  });
});

describe("sanitizeErrorMessage", () => {
  it("returns 'Unknown error' for non-Error values", () => {
    expect(sanitizeErrorMessage("string error")).toBe("Unknown error");
    expect(sanitizeErrorMessage(null)).toBe("Unknown error");
    expect(sanitizeErrorMessage(42)).toBe("Unknown error");
  });

  it("categorizes timeout errors", () => {
    expect(sanitizeErrorMessage(new Error("Connection timed out"))).toBe("Error: connection timed out");
    expect(sanitizeErrorMessage(new Error("Request timeout after 30s"))).toBe("Error: connection timed out");
  });

  it("categorizes authentication errors", () => {
    expect(sanitizeErrorMessage(new Error("Authentication failed for user@example.com"))).toBe(
      "Error: authentication failed",
    );
    expect(sanitizeErrorMessage(new Error("Invalid login credentials"))).toBe("Error: authentication failed");
  });

  it("categorizes connection errors with actionable detail", () => {
    // ECONNREFUSED → Bridge-not-running hint, with the host:port surfaced from
    // the error's structured fields (not sensitive — no credentials).
    const refused = Object.assign(new Error("connect ECONNREFUSED 127.0.0.1:1143"), {
      code: "ECONNREFUSED",
      address: "127.0.0.1",
      port: 1143,
    });
    const refusedMsg = sanitizeErrorMessage(refused);
    expect(refusedMsg).toContain("connection refused");
    expect(refusedMsg).toContain("127.0.0.1:1143");
    expect(refusedMsg).toContain("Proton Mail Bridge");

    // ENOTFOUND → host-not-found + config hint.
    const notFound = Object.assign(new Error("getaddrinfo ENOTFOUND smtp.example.com"), {
      code: "ENOTFOUND",
      address: "smtp.example.com",
    });
    const notFoundMsg = sanitizeErrorMessage(notFound);
    expect(notFoundMsg).toContain("host not found");
    expect(notFoundMsg).toContain("smtp.example.com");

    // Bare connection error with no structured fields still categorizes cleanly.
    expect(sanitizeErrorMessage(new Error("socket hang up while connecting"))).toContain("connection failed");
  });

  it("passes through 'not found' errors", () => {
    const err = new Error("Message UID 42 not found in INBOX");
    expect(sanitizeErrorMessage(err)).toBe("Error: Message UID 42 not found in INBOX");
  });

  it("passes through folder-not-found errors verbatim", () => {
    // Regression guard: translated IMAP errors like "Folder not found: X" must reach the user
    // in full so they can act on them rather than being generic-categorized.
    const folder = new Error("Folder not found: NonexistentFolder");
    expect(sanitizeErrorMessage(folder)).toBe("Error: Folder not found: NonexistentFolder");

    const dest = new Error("Destination folder not found: NonexistentFolder");
    expect(sanitizeErrorMessage(dest)).toBe("Error: Destination folder not found: NonexistentFolder");
  });

  it("redacts credentials in fallback messages", () => {
    const result = sanitizeErrorMessage(new Error('Unexpected: password="secret123"'));
    expect(result).not.toContain("secret123");
    expect(result).toContain("[REDACTED]");
  });

  it("redacts colon-separated credential forms in fallback messages", () => {
    // Messages containing the words "auth", "login", or "credential" are
    // categorized as auth-failures before redaction runs — those are tested
    // separately. Here we focus on the fallback path with `pass`/`password`/`user`.
    expect(sanitizeErrorMessage(new Error("unexpected pass=hunter2"))).toContain("[REDACTED]");
    expect(sanitizeErrorMessage(new Error("got user: alice@x.y"))).toContain("[REDACTED]");
    expect(sanitizeErrorMessage(new Error('SMTP error password="secret"'))).toContain("[REDACTED]");
  });

  it("does NOT redact prose that happens to start with a credential keyword", () => {
    // The previous regex matched `[=:\s]+` between keyword and value, so plain
    // English like "Pass either A or B" got mangled into "[REDACTED] A or B".
    // The fix requires an explicit `=` or `:` between keyword and value.
    const result = sanitizeErrorMessage(new Error("Pass either `markdownBody` OR `body`/`isHtml`, not both."));
    expect(result).not.toContain("[REDACTED]");
    expect(result).toContain("Pass either");
  });

  it("does NOT redact identifiers containing a credential keyword as a substring", () => {
    // "user_input" / "password_protected" shouldn't trigger — the keyword is
    // part of a larger identifier, not a key/value pair.
    const result = sanitizeErrorMessage(new Error("password_protected zip detected"));
    expect(result).not.toContain("[REDACTED]");
  });

  it("passes through messages up to the 500-char cap without truncation", () => {
    const msg = "A".repeat(450);
    const result = sanitizeErrorMessage(new Error(msg));
    // "Error: " (7) + 450 = 457
    expect(result.length).toBe(457);
    expect(result.endsWith("…")).toBe(false);
  });

  it("truncates messages longer than 500 chars at a word boundary with an ellipsis", () => {
    const msg = "Word ".repeat(120); // 600 chars, plenty of whitespace
    const result = sanitizeErrorMessage(new Error(msg));
    // "Error: " (7) + up to 500 (cap) + 1 ("…") = 508 max
    expect(result.length).toBeLessThanOrEqual(508);
    expect(result.endsWith("…")).toBe(true);
    // No partial word before the ellipsis (the truncation regex drops
    // trailing /\s\S*$/, so the char before … is either nothing or a word).
    const beforeEllipsis = result.slice(0, -1);
    expect(beforeEllipsis.endsWith(" ")).toBe(false);
  });
});

describe("validateBulkInput (XOR)", () => {
  it("accepts uids alone", () => {
    expect(() => validateBulkInput({ uids: [1, 2, 3] })).not.toThrow();
  });

  it("accepts match alone", () => {
    expect(() => validateBulkInput({ match: { from: "x@y.z" } })).not.toThrow();
  });

  it("rejects both uids and match with a differentiated (got both) message", () => {
    expect(() => validateBulkInput({ uids: [1], match: { from: "x" } })).toThrow(/got both/i);
  });

  it("rejects neither uids nor match with a differentiated (got neither) message", () => {
    expect(() => validateBulkInput({})).toThrow(/got neither/i);
  });

  it("rejects uids array exceeding 1000 entries", () => {
    const uids = Array.from({ length: 1001 }, (_, i) => i + 1);
    expect(() => validateBulkInput({ uids })).toThrow(/1000/);
  });

  it("rejects empty uids array", () => {
    expect(() => validateBulkInput({ uids: [] })).toThrow(/at least one/i);
  });
});

describe("permanentDeleteNeedsConfirm (bulk_delete expunge gate)", () => {
  it("requires confirmation for a live permanent delete", () => {
    expect(permanentDeleteNeedsConfirm({ permanent: true })).toBe(true);
    expect(permanentDeleteNeedsConfirm({ permanent: true, confirm: false })).toBe(true);
  });

  it("allows a live permanent delete once confirmed", () => {
    expect(permanentDeleteNeedsConfirm({ permanent: true, confirm: true })).toBe(false);
  });

  it("never gates a dry-run (preview destroys nothing), even permanent + unconfirmed", () => {
    expect(permanentDeleteNeedsConfirm({ permanent: true, dryRun: true })).toBe(false);
    expect(permanentDeleteNeedsConfirm({ permanent: true, confirm: false, dryRun: true })).toBe(false);
  });

  it("never gates a soft delete (goes to Trash, recoverable)", () => {
    expect(permanentDeleteNeedsConfirm({ permanent: false })).toBe(false);
    expect(permanentDeleteNeedsConfirm({})).toBe(false);
  });
});

describe("validateSizeBound (bytes)", () => {
  it("accepts positive integer below 1 GiB", () => {
    expect(validateSizeBound(1024)).toBe(1024);
    expect(validateSizeBound(1)).toBe(1);
  });

  it("rejects zero or negative", () => {
    expect(() => validateSizeBound(0)).toThrow();
    expect(() => validateSizeBound(-1)).toThrow();
  });

  it("rejects values above 1 GiB", () => {
    expect(() => validateSizeBound(1024 * 1024 * 1024 + 1)).toThrow(/1 GiB/);
  });

  it("rejects non-integers", () => {
    expect(() => validateSizeBound(1.5)).toThrow();
  });
});

describe("validateListId", () => {
  it("accepts a reasonable List-Id value", () => {
    expect(validateListId("politics.substack.com")).toBe("politics.substack.com");
  });

  it("rejects empty strings", () => {
    expect(() => validateListId("")).toThrow();
  });

  it("rejects strings longer than 200 chars", () => {
    expect(() => validateListId("x".repeat(201))).toThrow(/200/);
  });

  it("rejects strings containing control characters", () => {
    expect(() => validateListId("foo\x01bar")).toThrow();
  });
});

describe("resolveAllowlistedPath (download_attachment.saveTo safety)", () => {
  const allowDir = "/Users/test/downloads";
  // Identity realpath stub — by default, no symlink rewriting.
  const identityRealpath = async (p: import("node:fs").PathLike): Promise<string> => String(p);

  it("rejects when ALLOW_FILE_DOWNLOAD_DIR is not set", async () => {
    await expect(resolveAllowlistedPath("file.bin", undefined, identityRealpath)).rejects.toThrow(
      /requires ALLOW_FILE_DOWNLOAD_DIR/,
    );
  });

  it("rejects absolute saveTo paths", async () => {
    await expect(resolveAllowlistedPath("/etc/passwd", allowDir, identityRealpath)).rejects.toThrow(
      /must be a relative path/,
    );
  });

  it("rejects '..' traversal attempts that escape the allow root", async () => {
    await expect(resolveAllowlistedPath("../../../etc/passwd", allowDir, identityRealpath)).rejects.toThrow(
      /escapes ALLOW_FILE_DOWNLOAD_DIR/,
    );
  });

  it("rejects symlink escapes (parent realpath resolves outside the root)", async () => {
    // Mock realpath to claim the parent dir actually lives in /tmp — i.e. the
    // allow root contains a symlink pointing out of the allowlist.
    const escapingRealpath = async (p: import("node:fs").PathLike): Promise<string> => {
      const s = String(p);
      if (s === allowDir) return allowDir;
      return "/tmp/escape"; // parent dir of the saveTo candidate
    };
    await expect(resolveAllowlistedPath("subdir/file.bin", allowDir, escapingRealpath)).rejects.toThrow(/symlink/);
  });

  it("accepts a well-formed relative path inside the allow root", async () => {
    const result = await resolveAllowlistedPath("invoices/q1.pdf", allowDir, identityRealpath);
    expect(result).toBe(path.join(allowDir, "invoices", "q1.pdf"));
  });

  it("translates a missing ALLOW_FILE_DOWNLOAD_DIR into an actionable error (not raw ENOENT)", async () => {
    // UX: realpath on a non-existent configured dir threw a raw
    // `ENOENT ... realpath` at the caller. We now check the root first and emit
    // a create-it message. The root must be resolved BEFORE the parent dir so
    // this fires instead of a confusing parent-dir ENOENT.
    const enoentRealpath = async (p: import("node:fs").PathLike): Promise<string> => {
      const err = new Error(`ENOENT: no such file or directory, realpath '${String(p)}'`) as NodeJS.ErrnoException;
      err.code = "ENOENT";
      throw err;
    };
    await expect(resolveAllowlistedPath("file.bin", allowDir, enoentRealpath)).rejects.toThrow(
      /ALLOW_FILE_DOWNLOAD_DIR points to a directory that does not exist.*Create it first/s,
    );
  });

  it("translates a missing saveTo subdirectory into an actionable error", async () => {
    // Root exists, but the saveTo subdir doesn't → friendly message, not ENOENT.
    const rootOnlyRealpath = async (p: import("node:fs").PathLike): Promise<string> => {
      const s = String(p);
      if (s === allowDir) return allowDir;
      const err = new Error(`ENOENT: no such file or directory, realpath '${s}'`) as NodeJS.ErrnoException;
      err.code = "ENOENT";
      throw err;
    };
    await expect(resolveAllowlistedPath("nope/file.bin", allowDir, rootOnlyRealpath)).rejects.toThrow(
      /target directory for saveTo does not exist/,
    );
  });
});

describe("buildReplyAllRecipients (reply-all self-exclusion)", () => {
  const me = "me@example.com";

  it("normal case: sender stays primary To, original recipients become CC, self dropped", () => {
    const result = buildReplyAllRecipients(
      "Alice <alice@x.y>",
      "me@example.com, Bob <bob@x.y>",
      "Carol <carol@x.y>",
      me,
    );
    expect(result).toEqual({ to: "Alice <alice@x.y>", cc: "Bob <bob@x.y>, Carol <carol@x.y>" });
  });

  it("self-loop fix: reply-all to your own sent message targets the original recipients, not you", () => {
    // From is self → must be dropped from the primary To, not looped back.
    const result = buildReplyAllRecipients("me@example.com", "Bob <bob@x.y>", "Carol <carol@x.y>", me);
    expect(result).toEqual({ to: "Bob <bob@x.y>", cc: "Carol <carol@x.y>" });
    // self never appears anywhere
    expect(JSON.stringify(result)).not.toContain("me@example.com");
  });

  it("folds in and dedupes extra CC, excluding self", () => {
    const result = buildReplyAllRecipients("Alice <alice@x.y>", "bob@x.y", "", me, "dave@x.y, me@example.com, bob@x.y");
    // alice primary; bob (from To) + dave (extra) in CC; me excluded; bob not duplicated.
    expect(result).toEqual({ to: "Alice <alice@x.y>", cc: "bob@x.y, dave@x.y" });
  });

  it("dedupes the sender out of CC when they also appear in the original To (old code double-listed)", () => {
    // The previous buildReplyAllCc never deduped the primary target against the
    // CC set, so a sender who was also in their own To landed in BOTH To and CC.
    // The unified helper dedupes by address across the whole set.
    const result = buildReplyAllRecipients("Alice <alice@x.y>", "alice@x.y, Bob <bob@x.y>", "", me);
    expect(result).toEqual({ to: "Alice <alice@x.y>", cc: "Bob <bob@x.y>" });
  });

  it("returns null when nobody but self remains", () => {
    expect(buildReplyAllRecipients("me@example.com", "me@example.com", "", me)).toBeNull();
  });

  it("extractEmailAddress parses display-name and bare forms, lowercased", () => {
    expect(extractEmailAddress("Alice <Alice@X.Y>")).toBe("alice@x.y");
    expect(extractEmailAddress("BOB@x.y")).toBe("bob@x.y");
    expect(extractEmailAddress("not an address")).toBeNull();
  });
});

describe("externalRecipientAddresses (send-family outbound guard / preview)", () => {
  const me = "me@example.com";

  it("returns external addresses across To/Cc/Bcc, excluding self, deduped and lowercased", () => {
    const ext = externalRecipientAddresses(
      ["me@example.com, Alice <ALICE@x.y>", "bob@x.y", "alice@x.y, carol@x.y"],
      me,
    );
    expect(ext).toEqual(["alice@x.y", "bob@x.y", "carol@x.y"]);
  });

  it("returns [] when every recipient is the authenticated self (the self-only-OK case)", () => {
    expect(externalRecipientAddresses(["me@example.com", "Me <ME@example.com>"], me)).toEqual([]);
  });

  it("treats the self address angle-bracket form as self too", () => {
    expect(externalRecipientAddresses(["My Name <me@example.com>"], "My Name <me@example.com>")).toEqual([]);
  });

  it("ignores undefined/empty lists", () => {
    expect(externalRecipientAddresses([undefined, "", "  "], me)).toEqual([]);
  });

  it("fail-closed: a fragment with no parseable address is surfaced as external, not dropped", () => {
    // A malformed recipient must not silently bypass the guard — better to flag
    // it than let it through unseen.
    expect(externalRecipientAddresses(["garbled-no-at-sign"], me)).toEqual(["garbled-no-at-sign"]);
  });
});

describe("validateSubject", () => {
  it("accepts ordinary subjects", () => {
    expect(validateSubject("Hello, world")).toBe(true);
    expect(validateSubject("Re: [Project] update")).toBe(true);
    expect(validateSubject("Subject with unicode — émojis 🎉")).toBe(true);
  });

  it("accepts tab characters (used in some header folding contexts)", () => {
    expect(validateSubject("Tabby\tsubject")).toBe(true);
  });

  it("rejects subjects containing CR or LF (header-injection attempts)", () => {
    expect(validateSubject("Subject with\r\nBcc: smuggled@example.com")).toBe(false);
    expect(validateSubject("Subject with\rcarriage return")).toBe(false);
    expect(validateSubject("Subject with\nline feed")).toBe(false);
  });

  it("rejects subjects containing NUL or DEL", () => {
    expect(validateSubject("Subject with\x00null")).toBe(false);
    expect(validateSubject("Subject with\x7fdel")).toBe(false);
  });

  it("rejects other C0 control characters", () => {
    expect(validateSubject("Subject with\x01control")).toBe(false);
    expect(validateSubject("Subject with\x1fcontrol")).toBe(false);
  });
});

describe("sanitizeEmailHtml (email send/reply/forward HTML allowlist)", () => {
  it("strips <script> entirely", () => {
    const result = sanitizeEmailHtml("<p>Hi</p><script>alert(1)</script>");
    expect(result).not.toContain("script");
    expect(result).not.toContain("alert(1)");
    expect(result).toContain("Hi");
  });

  it("strips inline event handlers like onclick / onerror", () => {
    const result = sanitizeEmailHtml('<a href="https://example.com" onclick="evil()">click</a>');
    expect(result).not.toContain("onclick");
    expect(result).not.toContain("evil()");
    expect(result).toContain("https://example.com");
  });

  it("strips inline style attributes", () => {
    const result = sanitizeEmailHtml('<div style="position:fixed;top:0">overlay</div>');
    expect(result).not.toContain("style");
    expect(result).not.toContain("position:fixed");
    expect(result).toContain("overlay");
  });

  it("strips remote <img> beacons (http/https)", () => {
    const result = sanitizeEmailHtml('<img src="https://tracker.example.com/pixel.gif">');
    // sanitize-html keeps the tag but drops the disallowed src scheme — net
    // result is a <img> with no src, which is harmless. We test that the URL
    // itself is gone.
    expect(result).not.toContain("tracker.example.com");
  });

  it("allows cid: inline image references", () => {
    const result = sanitizeEmailHtml('<img src="cid:logo123" alt="Logo">');
    expect(result).toContain("cid:logo123");
    expect(result).toContain('alt="Logo"');
  });

  it("allows data: inline images (already embedded base64)", () => {
    const result = sanitizeEmailHtml('<img src="data:image/png;base64,iVBORw0KGgo=">');
    expect(result).toContain("data:image/png;base64");
  });

  it("preserves text-formatting and link tags", () => {
    const result = sanitizeEmailHtml('<p>Hello <b>world</b>!</p><a href="https://example.com">link</a>');
    expect(result).toContain("<p>");
    expect(result).toContain("<b>");
    expect(result).toContain('<a href="https://example.com">');
    expect(result).toContain("link");
  });

  it("strips <iframe> entirely", () => {
    const result = sanitizeEmailHtml('<iframe src="https://attacker.example.com"></iframe>');
    expect(result).not.toContain("iframe");
    expect(result).not.toContain("attacker.example.com");
  });
});

describe("detectActiveHtml (read_message preferHtml active-content flag)", () => {
  it("flags <script> blocks", () => {
    expect(detectActiveHtml("<p>hi</p><script>alert('xss')</script>")).toBe(true);
  });

  it("flags javascript: URLs", () => {
    expect(detectActiveHtml('<a href="javascript:void(0)">click</a>')).toBe(true);
  });

  it("flags inline event handlers (onerror/onload/onclick)", () => {
    expect(detectActiveHtml('<img src="x" onerror="steal()">')).toBe(true);
    expect(detectActiveHtml('<body onload="go()">')).toBe(true);
    expect(detectActiveHtml('<div onclick="x">y</div>')).toBe(true);
  });

  it("flags remote <img> beacons (http/https tracking pixels)", () => {
    expect(detectActiveHtml('<img src="http://evil.example.com/beacon.png">')).toBe(true);
    expect(detectActiveHtml('<img alt="x" src="https://tracker.test/p.gif" width="1">')).toBe(true);
  });

  it("flags <iframe>/<object>/<embed> frames", () => {
    expect(detectActiveHtml('<iframe src="https://attacker.test"></iframe>')).toBe(true);
    expect(detectActiveHtml('<object data="x"></object>')).toBe(true);
    expect(detectActiveHtml("<embed>")).toBe(true);
  });

  it("does NOT flag benign formatted HTML", () => {
    expect(detectActiveHtml("<p>Hello <b>world</b></p><a href='https://example.com'>link</a>")).toBe(false);
  });

  it("does NOT flag cid:/data: inline images (they don't phone home)", () => {
    expect(detectActiveHtml('<img src="cid:logo123" alt="Logo">')).toBe(false);
    expect(detectActiveHtml('<img src="data:image/png;base64,iVBORw0KGgo=">')).toBe(false);
  });

  it("returns false for empty/plain input", () => {
    expect(detectActiveHtml("")).toBe(false);
    expect(detectActiveHtml("just some plain text, no tags")).toBe(false);
  });
});

describe("snippetNote (includeSnippet untrusted-content banner)", () => {
  it("returns the untrusted banner when includeSnippet and at least one row has a snippet", () => {
    expect(snippetNote(true, [{ snippet: "hello" }, {}])).toBe(SNIPPET_UNTRUSTED_NOTE);
  });

  it("returns empty when includeSnippet is false/undefined (even if rows carry snippets)", () => {
    expect(snippetNote(false, [{ snippet: "hello" }])).toBe("");
    expect(snippetNote(undefined, [{ snippet: "hello" }])).toBe("");
  });

  it("returns empty when no row carries a snippet (nothing untrusted to flag)", () => {
    expect(snippetNote(true, [{}, { snippet: "" }])).toBe("");
    expect(snippetNote(true, [])).toBe("");
  });

  it("banner ends with a newline so it sits as a block before the rows, not inline with them", () => {
    // Placement contract: the caller puts this BEFORE the rows; the trailing
    // newline keeps the first row from being glued onto the banner text.
    expect(SNIPPET_UNTRUSTED_NOTE.endsWith("\n")).toBe(true);
    expect(SNIPPET_UNTRUSTED_NOTE).toContain("UNTRUSTED");
  });
});

describe("isValidBase64", () => {
  it("accepts standard base64 content", () => {
    expect(isValidBase64("aGVsbG8gd29ybGQ=")).toBe(true); // "hello world"
    expect(isValidBase64("Zm9v")).toBe(true); // "foo"
    expect(isValidBase64("Zm9vYg==")).toBe(true); // "foob"
  });

  it("accepts whitespace-wrapped base64 (line-wrapped output)", () => {
    expect(isValidBase64("aGVs\nbG8g\nd29y\nbGQ=")).toBe(true);
    expect(isValidBase64("  Zm9v  ")).toBe(true);
  });

  it("accepts URL-safe base64 alphabet (- and _ instead of + and /)", () => {
    expect(isValidBase64("aGVsbG8td29ybGQ=")).toBe(true);
    // 8 chars, padded, URL-safe alphabet
    expect(isValidBase64("Zm9vYmFy")).toBe(true);
  });

  it("rejects empty strings", () => {
    expect(isValidBase64("")).toBe(false);
    expect(isValidBase64("   ")).toBe(false);
  });

  it("rejects a '###not-base64###' payload", () => {
    expect(isValidBase64("###not-base64###")).toBe(false);
  });

  it("rejects content with invalid characters", () => {
    expect(isValidBase64("not base64!")).toBe(false);
    expect(isValidBase64("aGVsbG8=&extra")).toBe(false);
  });

  it("rejects content with wrong padding length (not a multiple of 4)", () => {
    expect(isValidBase64("aGVsbA")).toBe(false); // 6 chars without padding
    expect(isValidBase64("aGVsbG8gd29ybGQ")).toBe(false); // 15 chars, needs 1 more
  });
});

describe("isValidMimeContentType", () => {
  it("accepts canonical type/subtype values", () => {
    expect(isValidMimeContentType("application/pdf")).toBe(true);
    expect(isValidMimeContentType("image/png")).toBe(true);
    expect(isValidMimeContentType("text/plain")).toBe(true);
    expect(isValidMimeContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document")).toBe(
      true,
    );
  });

  it("accepts values with parameters", () => {
    expect(isValidMimeContentType("text/plain; charset=utf-8")).toBe(true);
    expect(isValidMimeContentType('application/json; charset="utf-8"')).toBe(true);
  });

  it("rejects a 'this/is/not/valid' payload", () => {
    expect(isValidMimeContentType("this/is/not/valid")).toBe(false);
  });

  it("rejects malformed types", () => {
    expect(isValidMimeContentType("")).toBe(false);
    expect(isValidMimeContentType("no-slash")).toBe(false);
    expect(isValidMimeContentType("application/")).toBe(false);
    expect(isValidMimeContentType("/png")).toBe(false);
    expect(isValidMimeContentType("application/pdf; ; charset=utf-8")).toBe(false);
    expect(isValidMimeContentType("application/pdf\r\ninjected: header")).toBe(false);
  });
});

import { assertSinceBeforeNotIdentical, looksLikeAddressInDisplayName } from "../validation.js";

describe("assertSinceBeforeNotIdentical", () => {
  it("throws when since equals before (IMAP BEFORE is exclusive — matches nothing)", () => {
    expect(() => assertSinceBeforeNotIdentical("2026-05-26", "2026-05-26")).toThrow(
      /equals before.*BEFORE is exclusive.*matches no messages/,
    );
  });

  it("error message names the actual date so callers can grep / regex it", () => {
    try {
      assertSinceBeforeNotIdentical("2026-05-26", "2026-05-26");
      throw new Error("should have thrown");
    } catch (err) {
      expect((err as Error).message).toContain("2026-05-26");
      expect((err as Error).message).toMatch(/one day later/);
    }
  });

  it("passes when since < before", () => {
    expect(() => assertSinceBeforeNotIdentical("2026-05-01", "2026-05-26")).not.toThrow();
  });

  it("passes when only since is set", () => {
    expect(() => assertSinceBeforeNotIdentical("2026-05-26", undefined)).not.toThrow();
  });

  it("passes when only before is set", () => {
    expect(() => assertSinceBeforeNotIdentical(undefined, "2026-05-26")).not.toThrow();
  });

  it("passes when neither is set", () => {
    expect(() => assertSinceBeforeNotIdentical(undefined, undefined)).not.toThrow();
  });

  it("does NOT throw when since > before (semantically empty but a different bug than same-day)", () => {
    // since > before is silently empty but represents a different caller error
    // (mixed-up order). We only catch the same-day case here; the > case
    // remains correctly empty and falls out as zero results.
    expect(() => assertSinceBeforeNotIdentical("2026-06-01", "2026-05-01")).not.toThrow();
  });
});

describe("looksLikeAddressInDisplayName", () => {
  it("returns true for a display-name spoofing payload", () => {
    expect(looksLikeAddressInDisplayName("Anthropic Security <security@anthropic.com>")).toBe(true);
    // After sanitizeFromName strips the brackets:
    expect(looksLikeAddressInDisplayName("Anthropic Security security@anthropic.com")).toBe(true);
  });

  it("returns true for any value containing @", () => {
    expect(looksLikeAddressInDisplayName("@John's iPhone")).toBe(true);
    expect(looksLikeAddressInDisplayName("user@example.com")).toBe(true);
    expect(looksLikeAddressInDisplayName("Acme @ HQ")).toBe(true);
  });

  it("returns false for ordinary display names", () => {
    expect(looksLikeAddressInDisplayName("Anthropic Security")).toBe(false);
    expect(looksLikeAddressInDisplayName("CEO of Acme")).toBe(false);
    expect(looksLikeAddressInDisplayName("John Doe")).toBe(false);
    expect(looksLikeAddressInDisplayName("O'Brien & Co.")).toBe(false);
    expect(looksLikeAddressInDisplayName("")).toBe(false);
  });

  it("returns false for full-width @ (U+FF20) — only ASCII @ is treated as the spoofing signal", () => {
    // Documenting current behavior: the regex matches plain ASCII @ only.
    // Full-width @ is a Unicode-confusable that mail clients will likely
    // render as a literal character, not as part of an address-like token.
    // If this becomes a real attack vector, tighten here with a NFKC normalization.
    expect(looksLikeAddressInDisplayName("Spoofer＠example.com")).toBe(false);
  });
});
