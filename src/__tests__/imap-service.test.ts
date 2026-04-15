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
    };
  }),
}));

const baseConfig: ImapConfig = {
  host: "127.0.0.1",
  port: 1143,
  secure: false,
  auth: { user: "test@pm.me", pass: "test-bridge-password" },
};

describe("ImapService", () => {
  beforeEach(() => {
    vi.clearAllMocks();
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

      mockFetch.mockReturnValueOnce((async function* () {
        for (const m of messages) yield m;
      })());

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
        bodyParts: new Map([["TEXT", Buffer.from("<html><body><p>Hello</p></body></html>")]]),
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 43);

      expect(msg.text).toBe("");
      expect(msg.html).toBe("<html><body><p>Hello</p></body></html>");
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

      mockFetch.mockReturnValueOnce((async function* () {
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
      })());

      const service = new ImapService(baseConfig);
      const result = await service.searchMessages("INBOX", { from: "alice@example.com" }, 10);

      expect(result).toHaveLength(2);
      expect(mockSearch).toHaveBeenCalledWith(
        expect.objectContaining({ from: "alice@example.com" }),
        { uid: true },
      );
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
});
