import { describe, it, expect, vi, beforeEach } from "vitest";
import { ImapService } from "../imap-service.js";
import type { ImapConfig } from "../imap-service.js";

// Mock imapflow
const mockConnect = vi.fn().mockResolvedValue(undefined);
const mockLogout = vi.fn().mockResolvedValue(undefined);
const mockList = vi.fn();
const mockStatus = vi.fn();
const mockFetch = vi.fn();
const mockFetchOne = vi.fn();
const mockSearch = vi.fn();
const mockGetMailboxLock = vi.fn().mockResolvedValue({ release: vi.fn() });
const mockMessageMove = vi.fn().mockResolvedValue({ uidMap: new Map() });
const mockMessageDelete = vi.fn().mockResolvedValue(true);
const mockMessageFlagsAdd = vi.fn().mockResolvedValue(true);
const mockMessageFlagsRemove = vi.fn().mockResolvedValue(true);
const mockDownload = vi.fn();

vi.mock("imapflow", () => ({
  ImapFlow: vi.fn().mockImplementation(function () {
    return {
      connect: mockConnect,
      logout: mockLogout,
      list: mockList,
      status: mockStatus,
      fetch: mockFetch,
      fetchOne: mockFetchOne,
      search: mockSearch,
      getMailboxLock: mockGetMailboxLock,
      messageMove: mockMessageMove,
      messageDelete: mockMessageDelete,
      messageFlagsAdd: mockMessageFlagsAdd,
      messageFlagsRemove: mockMessageFlagsRemove,
      download: mockDownload,
    };
  }),
}));

const baseConfig: ImapConfig = {
  host: "127.0.0.1",
  port: 1143,
  secure: false,
  auth: { user: "test@pm.me", pass: "test-bridge-password" },
};

const debugConfig: ImapConfig = { ...baseConfig, debug: true };

describe("ImapService", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("debug logging", () => {
    it("logs to stderr when debug is true", async () => {
      const spy = vi.spyOn(console, "error").mockImplementation(() => {});
      mockList.mockResolvedValueOnce([]);

      const service = new ImapService(debugConfig);
      await service.listFolders();

      expect(spy).toHaveBeenCalledWith("[IMAP] Listing folders");
      spy.mockRestore();
    });
  });

  describe("listFolders", () => {
    it("returns formatted folder list with message counts", async () => {
      mockList.mockResolvedValueOnce([
        { path: "INBOX", name: "Inbox", specialUse: "\\Inbox", status: { messages: 42, unseen: 3 } },
        { path: "Sent", name: "Sent", specialUse: "\\Sent", status: { messages: 100, unseen: 0 } },
        { path: "Archive", name: "Archive", specialUse: "", status: { messages: 500, unseen: 0 } },
      ]);

      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();

      expect(folders).toHaveLength(3);
      expect(folders[0]).toEqual({
        path: "INBOX",
        name: "Inbox",
        specialUse: "\\Inbox",
        messages: 42,
        unseen: 3,
      });
      expect(mockConnect).toHaveBeenCalled();
      expect(mockLogout).toHaveBeenCalled();
    });

    it("handles folders with missing status fields", async () => {
      mockList.mockResolvedValueOnce([
        { path: "INBOX", name: "Inbox", specialUse: "", status: undefined },
        { path: "Drafts", name: "Drafts", specialUse: undefined, status: { messages: undefined, unseen: undefined } },
      ]);

      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();

      expect(folders[0]).toEqual({
        path: "INBOX",
        name: "Inbox",
        specialUse: "",
        messages: 0,
        unseen: 0,
      });
      expect(folders[1]).toEqual({
        path: "Drafts",
        name: "Drafts",
        specialUse: "",
        messages: 0,
        unseen: 0,
      });
    });

    it("handles empty mailbox list", async () => {
      mockList.mockResolvedValueOnce([]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders).toEqual([]);
    });
  });

  describe("listMessages", () => {
    it("returns messages in reverse order (newest first)", async () => {
      mockStatus.mockResolvedValueOnce({ messages: 50 });

      // Simulate async generator
      const messages = [
        {
          uid: 48,
          envelope: {
            subject: "Older message",
            from: [{ name: "Alice", address: "alice@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-10T10:00:00Z"),
          },
          flags: new Set(["\\Seen"]),
        },
        {
          uid: 50,
          envelope: {
            subject: "Newer message",
            from: [{ address: "bob@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-15T12:00:00Z"),
          },
          flags: new Set(),
        },
      ];

      mockFetch.mockReturnValueOnce(
        (async function* () {
          for (const m of messages) yield m;
        })(),
      );

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 10);

      expect(result).toHaveLength(2);
      // Newest first
      expect(result[0].uid).toBe(50);
      expect(result[0].from).toBe("bob@example.com");
      expect(result[1].uid).toBe(48);
      expect(result[1].from).toBe("Alice <alice@example.com>");
      expect(result[1].flags).toEqual(["\\Seen"]);
    });

    it("returns empty array for empty folder", async () => {
      mockStatus.mockResolvedValueOnce({ messages: 0 });

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 10);
      expect(result).toEqual([]);
    });

    it("handles messages with missing envelope fields", async () => {
      mockStatus.mockResolvedValueOnce({ messages: 1 });

      mockFetch.mockReturnValueOnce(
        (async function* () {
          yield {
            uid: 1,
            envelope: {},
            flags: undefined,
          };
        })(),
      );

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 10);

      expect(result).toHaveLength(1);
      expect(result[0].subject).toBe("");
      expect(result[0].from).toBe("");
      expect(result[0].to).toBe("");
      expect(result[0].date).toBe("");
      expect(result[0].flags).toEqual([]);
    });

    it("paginates with beforeUid parameter", async () => {
      mockSearch.mockResolvedValueOnce([5, 10, 15, 20, 25]);

      const messages = [
        {
          uid: 20,
          envelope: {
            subject: "Older",
            from: [{ address: "alice@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-10T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 25,
          envelope: {
            subject: "Less old",
            from: [{ address: "bob@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-11T10:00:00Z"),
          },
          flags: new Set(),
        },
      ];

      mockFetch.mockReturnValueOnce(
        (async function* () {
          for (const m of messages) yield m;
        })(),
      );

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 2, 30);

      expect(mockSearch).toHaveBeenCalledWith({ uid: "1:29" }, { uid: true });
      expect(result).toHaveLength(2);
      expect(result[0].uid).toBe(25);
      expect(result[1].uid).toBe(20);
    });

    it("returns empty array when beforeUid search finds no results", async () => {
      mockSearch.mockResolvedValueOnce([]);

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 10, 5);
      expect(result).toEqual([]);
    });
  });

  describe("readMessage", () => {
    it("reads plain text message", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        envelope: {
          subject: "Test Subject",
          from: [{ name: "Sender", address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          cc: [{ address: "cc@example.com" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<abc123@example.com>",
        },
        flags: new Set(["\\Seen"]),
        bodyStructure: { type: "text/plain", part: "1" },
        bodyParts: new Map([["TEXT", Buffer.from("Hello, this is the body.")]]),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 42);

      expect(msg.uid).toBe(42);
      expect(msg.subject).toBe("Test Subject");
      expect(msg.from).toBe("Sender <sender@example.com>");
      expect(msg.text).toBe("Hello, this is the body.");
      expect(msg.html).toBe("");
    });

    it("reads HTML message", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 43,
        envelope: {
          subject: "HTML Email",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<html123@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/alternative",
          childNodes: [
            { type: "text/plain", part: "1" },
            { type: "text/html", part: "2" },
          ],
        },
        bodyParts: new Map([["TEXT", Buffer.from("<html><body><p>Hello</p></body></html>")]]),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 43);

      expect(msg.text).toBe("");
      expect(msg.html).toBe("<html><body><p>Hello</p></body></html>");
    });

    it("extracts attachment metadata from bodyStructure", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 44,
        envelope: {
          subject: "With Attachment",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<att123@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/plain", part: "1" },
            {
              type: "application/pdf",
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "report.pdf" },
              size: 12345,
            },
          ],
        },
        bodyParts: new Map([["TEXT", Buffer.from("See attached.")]]),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 44);

      expect(msg.attachments).toHaveLength(1);
      expect(msg.attachments[0]).toEqual({
        partNumber: "2",
        filename: "report.pdf",
        contentType: "application/pdf",
        size: 12345,
      });
      expect(msg.text).toBe("See attached.");
    });

    it("falls back to bodyParts '1' when 'TEXT' is absent", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 47,
        envelope: {
          subject: "Fallback",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<fb@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/plain", part: "1" },
        bodyParts: new Map([["1", Buffer.from("From part 1")]]),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 47);
      expect(msg.text).toBe("From part 1");
    });

    it("handles missing envelope fields with defaults", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 45,
        envelope: {},
        flags: undefined,
        bodyStructure: undefined,
        bodyParts: new Map(),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 45);

      expect(msg.subject).toBe("");
      expect(msg.from).toBe("");
      expect(msg.to).toBe("");
      expect(msg.cc).toBe("");
      expect(msg.date).toBe("");
      expect(msg.messageId).toBe("");
      expect(msg.flags).toEqual([]);
      expect(msg.text).toBe("");
      expect(msg.html).toBe("");
      expect(msg.attachments).toEqual([]);
    });

    it("extracts inline non-text parts as attachments", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 46,
        envelope: {
          subject: "Inline Image",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<inline123@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/related",
          childNodes: [
            { type: "text/html", part: "1" },
            {
              type: "image/png",
              part: "2",
              disposition: "inline",
              parameters: { name: "logo.png" },
              size: 5000,
            },
          ],
        },
        bodyParts: new Map([["TEXT", Buffer.from("<img src='cid:logo'>")]]),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 46);

      expect(msg.attachments).toHaveLength(1);
      expect(msg.attachments[0]).toEqual({
        partNumber: "2",
        filename: "logo.png",
        contentType: "image/png",
        size: 5000,
      });
      expect(msg.html).toBe("<img src='cid:logo'>");
    });

    it("throws when message not found", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.readMessage("INBOX", 999)).rejects.toThrow("Message UID 999 not found in INBOX");
    });
  });

  describe("searchMessages", () => {
    it("searches by criteria and returns results", async () => {
      mockSearch.mockResolvedValueOnce([10, 20, 30]);

      mockFetch.mockReturnValueOnce(
        (async function* () {
          yield {
            uid: 30,
            envelope: {
              subject: "Match 3",
              from: [{ address: "alice@example.com" }],
              to: [{ address: "me@pm.me" }],
              date: new Date("2026-04-15T12:00:00Z"),
            },
            flags: new Set(),
          };
          yield {
            uid: 20,
            envelope: {
              subject: "Match 2",
              from: [{ address: "alice@example.com" }],
              to: [{ address: "me@pm.me" }],
              date: new Date("2026-04-14T12:00:00Z"),
            },
            flags: new Set(["\\Flagged"]),
          };
        })(),
      );

      const service = new ImapService(baseConfig);
      const result = await service.searchMessages("INBOX", { from: "alice@example.com" }, 10);

      expect(result).toHaveLength(2);
      expect(mockSearch).toHaveBeenCalledWith(expect.objectContaining({ from: "alice@example.com" }), { uid: true });
    });

    it("returns empty array when no matches", async () => {
      mockSearch.mockResolvedValueOnce([]);

      const service = new ImapService(baseConfig);
      const result = await service.searchMessages("INBOX", { subject: "nonexistent" }, 10);
      expect(result).toEqual([]);
    });

    it("returns empty array when search returns false", async () => {
      mockSearch.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      const result = await service.searchMessages("INBOX", { subject: "nonexistent" }, 10);
      expect(result).toEqual([]);
    });
  });

  describe("downloadAttachment", () => {
    it("downloads and returns base64-encoded content", async () => {
      const fileContent = Buffer.from("PDF file contents here");
      mockDownload.mockResolvedValueOnce({
        meta: {
          contentType: "application/pdf",
          filename: "report.pdf",
        },
        content: (async function* () {
          yield fileContent;
        })(),
      });

      const service = new ImapService(baseConfig);
      const result = await service.downloadAttachment("INBOX", 42, "2");

      expect(result.filename).toBe("report.pdf");
      expect(result.contentType).toBe("application/pdf");
      expect(Buffer.from(result.content, "base64").toString()).toBe("PDF file contents here");
      expect(mockDownload).toHaveBeenCalledWith("42", "2", { uid: true });
    });

    it("concatenates multiple stream chunks", async () => {
      mockDownload.mockResolvedValueOnce({
        meta: { contentType: "text/plain", filename: "big.txt" },
        content: (async function* () {
          yield Buffer.from("chunk1-");
          yield Buffer.from("chunk2-");
          yield Buffer.from("chunk3");
        })(),
      });

      const service = new ImapService(baseConfig);
      const result = await service.downloadAttachment("INBOX", 42, "1");

      expect(Buffer.from(result.content, "base64").toString()).toBe("chunk1-chunk2-chunk3");
    });

    it("handles missing metadata gracefully", async () => {
      mockDownload.mockResolvedValueOnce({
        meta: {},
        content: (async function* () {
          yield Buffer.from("data");
        })(),
      });

      const service = new ImapService(baseConfig);
      const result = await service.downloadAttachment("INBOX", 42, "3");

      expect(result.filename).toBe("unnamed");
      expect(result.contentType).toBe("application/octet-stream");
    });
  });

  describe("moveMessage", () => {
    it("moves message and returns true on success", async () => {
      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");

      expect(result).toBe(true);
      expect(mockMessageMove).toHaveBeenCalledWith("42", "Archive", { uid: true });
      expect(mockConnect).toHaveBeenCalled();
      expect(mockLogout).toHaveBeenCalled();
    });

    it("returns false when server returns false", async () => {
      mockMessageMove.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");
      expect(result).toBe(false);
    });
  });

  describe("deleteMessage", () => {
    it("deletes message and returns true on success", async () => {
      const service = new ImapService(baseConfig);
      const result = await service.deleteMessage("INBOX", 42);

      expect(result).toBe(true);
      expect(mockMessageDelete).toHaveBeenCalledWith("42", { uid: true });
      expect(mockConnect).toHaveBeenCalled();
      expect(mockLogout).toHaveBeenCalled();
    });

    it("returns false when server returns false", async () => {
      mockMessageDelete.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      const result = await service.deleteMessage("INBOX", 42);
      expect(result).toBe(false);
    });
  });

  describe("folder path validation", () => {
    it("rejects folder names with control characters", async () => {
      const service = new ImapService(baseConfig);
      await expect(service.listMessages("IN\r\nBOX", 10)).rejects.toThrow("invalid control characters");
      await expect(service.readMessage("IN\x00BOX", 1)).rejects.toThrow("invalid control characters");
      await expect(service.searchMessages("IN\r\nBOX", {}, 10)).rejects.toThrow("invalid control characters");
      await expect(service.deleteMessage("IN\x00BOX", 1)).rejects.toThrow("invalid control characters");
      await expect(service.updateFlags("IN\r\nBOX", 1, [], [])).rejects.toThrow("invalid control characters");
    });

    it("rejects invalid folder in moveMessage for both source and destination", async () => {
      const service = new ImapService(baseConfig);
      await expect(service.moveMessage("IN\r\nBOX", 1, "Archive")).rejects.toThrow("invalid control characters");
      await expect(service.moveMessage("INBOX", 1, "Arch\x00ive")).rejects.toThrow("invalid control characters");
    });

    it("rejects invalid folder in downloadAttachment", async () => {
      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("IN\r\nBOX", 1, "1")).rejects.toThrow("invalid control characters");
    });
  });

  describe("partNumber validation", () => {
    it("rejects invalid part numbers in downloadAttachment", async () => {
      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("INBOX", 1, "../../etc/passwd")).rejects.toThrow(
        "Invalid MIME part number",
      );
      await expect(service.downloadAttachment("INBOX", 1, "abc")).rejects.toThrow("Invalid MIME part number");
      await expect(service.downloadAttachment("INBOX", 1, "")).rejects.toThrow("Invalid MIME part number");
    });
  });

  describe("attachment size limit", () => {
    it("aborts download when attachment exceeds max size", async () => {
      const destroyFn = vi.fn();
      const largeChunk = Buffer.alloc(1024 * 1024); // 1MB chunks
      mockDownload.mockResolvedValueOnce({
        meta: { contentType: "application/octet-stream", filename: "huge.bin" },
        content: Object.assign(
          (async function* () {
            for (let i = 0; i < 30; i++) yield largeChunk; // 30MB total
          })(),
          { destroy: destroyFn },
        ),
      });

      const service = new ImapService({ ...baseConfig, maxAttachmentSize: 10 * 1024 * 1024 }); // 10MB limit
      await expect(service.downloadAttachment("INBOX", 42, "2")).rejects.toThrow("exceeds maximum size");
    });
  });

  describe("updateFlags", () => {
    it("adds flags when flagsToAdd is non-empty", async () => {
      const service = new ImapService(baseConfig);
      await service.updateFlags("INBOX", 42, ["\\Seen", "\\Flagged"], []);

      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("42", ["\\Seen", "\\Flagged"], { uid: true });
      expect(mockMessageFlagsRemove).not.toHaveBeenCalled();
    });

    it("removes flags when flagsToRemove is non-empty", async () => {
      const service = new ImapService(baseConfig);
      await service.updateFlags("INBOX", 42, [], ["\\Seen"]);

      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("42", ["\\Seen"], { uid: true });
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
    });

    it("adds and removes flags in one call", async () => {
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 42, ["\\Flagged"], ["\\Seen"]);

      expect(result).toBe(true);
      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("42", ["\\Flagged"], { uid: true });
      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("42", ["\\Seen"], { uid: true });
    });
  });
});
