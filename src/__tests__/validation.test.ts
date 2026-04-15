import { describe, it, expect } from "vitest";
import {
  validateFolderPath,
  validatePartNumber,
  sanitizeFilename,
  sanitizeFromName,
  isValidEmailAddress,
  validateImapFlag,
  sanitizeErrorMessage,
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

  it("categorizes connection errors", () => {
    expect(sanitizeErrorMessage(new Error("connect ECONNREFUSED 127.0.0.1:1143"))).toBe("Error: connection failed");
    expect(sanitizeErrorMessage(new Error("getaddrinfo ENOTFOUND smtp.example.com"))).toBe("Error: connection failed");
  });

  it("passes through 'not found' errors", () => {
    const err = new Error("Message UID 42 not found in INBOX");
    expect(sanitizeErrorMessage(err)).toBe("Error: Message UID 42 not found in INBOX");
  });

  it("redacts credentials in fallback messages", () => {
    const result = sanitizeErrorMessage(new Error('Unexpected: password="secret123"'));
    expect(result).not.toContain("secret123");
    expect(result).toContain("[REDACTED]");
  });

  it("truncates long error messages", () => {
    const longMsg = "A".repeat(300);
    const result = sanitizeErrorMessage(new Error(longMsg));
    expect(result.length).toBeLessThanOrEqual(208); // "Error: " + 200 + null-ish
  });
});
