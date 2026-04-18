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
const mockAppend = vi.fn();

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
      append: mockAppend,
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
      mockSearch.mockResolvedValueOnce([48, 50]);

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

      expect(mockSearch).toHaveBeenCalledWith({ all: true }, { uid: true });
      expect(result).toHaveLength(2);
      // Newest first
      expect(result[0].uid).toBe(50);
      expect(result[0].from).toBe("bob@example.com");
      expect(result[1].uid).toBe(48);
      expect(result[1].from).toBe("Alice <alice@example.com>");
      expect(result[1].flags).toEqual(["\\Seen"]);
    });

    it("returns empty array for empty folder", async () => {
      mockSearch.mockResolvedValueOnce([]);

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 10);
      expect(result).toEqual([]);
    });

    it("handles messages with missing envelope fields", async () => {
      mockSearch.mockResolvedValueOnce([1]);

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

    it("sorts by date even when UIDs arrive out of order", async () => {
      mockSearch.mockResolvedValueOnce([100, 110, 121]);

      const messages = [
        {
          uid: 121,
          envelope: {
            subject: "Middle by date",
            from: [{ address: "a@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-12T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 100,
          envelope: {
            subject: "Newest by date",
            from: [{ address: "b@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-15T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 110,
          envelope: {
            subject: "Oldest by date",
            from: [{ address: "c@example.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-10T10:00:00Z"),
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

      expect(result[0].subject).toBe("Newest by date");
      expect(result[1].subject).toBe("Middle by date");
      expect(result[2].subject).toBe("Oldest by date");
    });

    it("returns empty array when beforeUid search finds no results", async () => {
      mockSearch.mockResolvedValueOnce([]);

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 10, 5);
      expect(result).toEqual([]);
    });

    it("correctly selects newest-by-date even when old UIDs have newer dates (moved messages)", async () => {
      // Scenario: UIDs 5,10 are old messages that were moved INTO the folder (so low UIDs).
      // UIDs 115,116,117,118,119,120 are newer by UID but OLDER by date.
      // With limit=5, the result must include UIDs 5 and 10 because they have the newest dates.
      mockSearch.mockResolvedValueOnce([5, 10, 115, 116, 117, 118, 119, 120]);

      const messages = [
        // Low UIDs but RECENT dates (moved into folder)
        {
          uid: 5,
          envelope: {
            subject: "Moved msg 1",
            from: [{ address: "a@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-15T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 10,
          envelope: {
            subject: "Moved msg 2",
            from: [{ address: "b@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-14T10:00:00Z"),
          },
          flags: new Set(),
        },
        // High UIDs but OLDER dates (original messages)
        {
          uid: 115,
          envelope: {
            subject: "Old msg 1",
            from: [{ address: "c@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-03-01T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 116,
          envelope: {
            subject: "Old msg 2",
            from: [{ address: "d@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-03-02T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 117,
          envelope: {
            subject: "Old msg 3",
            from: [{ address: "e@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-03-03T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 118,
          envelope: {
            subject: "Old msg 4",
            from: [{ address: "f@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-03-04T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 119,
          envelope: {
            subject: "Old msg 5",
            from: [{ address: "g@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-03-05T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 120,
          envelope: {
            subject: "Old msg 6",
            from: [{ address: "h@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-03-06T10:00:00Z"),
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
      const result = await service.listMessages("INBOX", 5);

      // The two moved messages should appear first (newest by date)
      expect(result).toHaveLength(5);
      expect(result[0].uid).toBe(5); // Apr 15 — newest
      expect(result[0].subject).toBe("Moved msg 1");
      expect(result[1].uid).toBe(10); // Apr 14
      expect(result[1].subject).toBe("Moved msg 2");
      // Followed by the 3 most recent of the old messages
      expect(result[2].uid).toBe(120); // Mar 6
      expect(result[3].uid).toBe(119); // Mar 5
      expect(result[4].uid).toBe(118); // Mar 4
    });

    it("fetches ALL UIDs when total <= 500 for exact date ordering", async () => {
      // Even with limit=2, all UIDs should be fetched so the sort is exact
      mockSearch.mockResolvedValueOnce([1, 2, 3, 4, 5]);

      const messages = [
        {
          uid: 1,
          envelope: {
            subject: "Newest",
            from: [{ address: "a@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-15T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 2,
          envelope: {
            subject: "Middle",
            from: [{ address: "b@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-10T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 3,
          envelope: {
            subject: "Oldest",
            from: [{ address: "c@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-05T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 4,
          envelope: {
            subject: "Second",
            from: [{ address: "d@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-14T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 5,
          envelope: {
            subject: "Third",
            from: [{ address: "e@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-09T10:00:00Z"),
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
      const result = await service.listMessages("INBOX", 2);

      // All 5 UIDs should be fetched (not just 2 or 10)
      expect(mockFetch).toHaveBeenCalledWith("1,2,3,4,5", expect.anything(), { uid: true });
      // Top 2 by date
      expect(result).toHaveLength(2);
      expect(result[0].subject).toBe("Newest"); // Apr 15
      expect(result[1].subject).toBe("Second"); // Apr 14
    });
  });

  describe("searchMessages ordering", () => {
    it("returns search results sorted by date, not UID order", async () => {
      // Simulate the exact bug: search returns UIDs in ascending order,
      // but dates don't correlate with UIDs
      mockSearch.mockResolvedValueOnce([69, 70, 71, 72, 73, 112, 113, 114, 115, 121]);

      const messages = [
        {
          uid: 121,
          envelope: {
            subject: "Most recent",
            from: [{ address: "a@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-15T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 112,
          envelope: {
            subject: "2nd most recent",
            from: [{ address: "b@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-14T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 113,
          envelope: {
            subject: "3rd",
            from: [{ address: "c@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-13T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 114,
          envelope: {
            subject: "4th",
            from: [{ address: "d@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-12T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 115,
          envelope: {
            subject: "5th",
            from: [{ address: "e@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-04-11T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 69,
          envelope: {
            subject: "Old 1",
            from: [{ address: "f@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-01-10T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 70,
          envelope: {
            subject: "Old 2",
            from: [{ address: "g@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-01-11T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 71,
          envelope: {
            subject: "Old 3",
            from: [{ address: "h@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-01-12T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 72,
          envelope: {
            subject: "Old 4",
            from: [{ address: "i@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-01-13T10:00:00Z"),
          },
          flags: new Set(),
        },
        {
          uid: 73,
          envelope: {
            subject: "Old 5",
            from: [{ address: "j@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-01-14T10:00:00Z"),
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
      const result = await service.searchMessages("INBOX", { since: "2026-01-01" }, 5);

      // Must return the 5 newest BY DATE, not by UID
      expect(result).toHaveLength(5);
      expect(result[0].subject).toBe("Most recent"); // UID 121, Apr 15
      expect(result[1].subject).toBe("2nd most recent"); // UID 112, Apr 14
      expect(result[2].subject).toBe("3rd"); // UID 113, Apr 13
      expect(result[3].subject).toBe("4th"); // UID 114, Apr 12
      expect(result[4].subject).toBe("5th"); // UID 115, Apr 11
      // UIDs 69-73 (oldest by date) should NOT appear
    });

    it("fetches all matching UIDs for correct date ordering (not just last N)", async () => {
      // All 10 UIDs should be fetched, not just the last 5
      mockSearch.mockResolvedValueOnce([69, 70, 71, 72, 73, 112, 113, 114, 115, 121]);

      const messages = [
        {
          uid: 69,
          envelope: {
            subject: "A",
            from: [{ address: "a@x.com" }],
            to: [{ address: "me@pm.me" }],
            date: new Date("2026-01-01T10:00:00Z"),
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
      await service.searchMessages("INBOX", { from: "test" }, 5);

      // Verify ALL 10 UIDs were fetched, not just the last 5
      expect(mockFetch).toHaveBeenCalledWith("69,70,71,72,73,112,113,114,115,121", expect.anything(), { uid: true });
    });
  });

  describe("readMessage", () => {
    function mockDownloadPart(text: string) {
      mockDownload.mockResolvedValueOnce({
        meta: {},
        content: (async function* () {
          yield Buffer.from(text);
        })(),
      });
    }

    it("reads plain text message via download()", async () => {
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
      });
      mockDownloadPart("Hello, this is the body.");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 42);

      expect(msg.uid).toBe(42);
      expect(msg.subject).toBe("Test Subject");
      expect(msg.from).toBe("Sender <sender@example.com>");
      expect(msg.body).toBe("Hello, this is the body.");
      expect(msg.bodyFormat).toBe("text");
      expect(msg.truncated).toBe(false);
      expect(msg.originalLength).toBeUndefined();
      expect(mockDownload).toHaveBeenCalledWith("42", "1", { uid: true });
    });

    it("prefers text/plain over HTML in multipart/alternative", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 43,
        envelope: {
          subject: "Multi",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<multi@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/alternative",
          childNodes: [
            { type: "text/plain", part: "1" },
            { type: "text/html", part: "2" },
          ],
        },
      });
      mockDownloadPart("Plain text version");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 43);

      expect(msg.body).toBe("Plain text version");
      expect(msg.bodyFormat).toBe("text");
      // Should download part "1" (text/plain), not "2" (text/html)
      expect(mockDownload).toHaveBeenCalledWith("43", "1", { uid: true });
    });

    it("strips HTML when only HTML part exists", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 44,
        envelope: {
          subject: "HTML Only",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<html@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/html", part: "1" },
      });
      mockDownloadPart("<html><body><p>Hello</p><p>World</p></body></html>");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 44);

      expect(msg.bodyFormat).toBe("html-stripped");
      expect(msg.body).not.toContain("<");
      expect(msg.body).toContain("Hello");
      expect(msg.body).toContain("World");
    });

    it("returns raw HTML when preferHtml is true", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 45,
        envelope: {
          subject: "HTML Pref",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<htmlpref@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/alternative",
          childNodes: [
            { type: "text/plain", part: "1" },
            { type: "text/html", part: "2" },
          ],
        },
      });
      mockDownloadPart("<p>Raw HTML</p>");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 45, { preferHtml: true });

      expect(msg.body).toBe("<p>Raw HTML</p>");
      expect(msg.bodyFormat).toBe("html");
      expect(mockDownload).toHaveBeenCalledWith("45", "2", { uid: true });
    });

    it("truncates body to maxBodyLength", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 46,
        envelope: {
          subject: "Long",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<long@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/plain", part: "1" },
      });
      mockDownloadPart("A".repeat(1000));

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 46, { maxBodyLength: 100 });

      expect(msg.body).toHaveLength(100);
      expect(msg.truncated).toBe(true);
      expect(msg.originalLength).toBe(1000);
    });

    it("reads single-part text/plain with no part field on root structure", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 50,
        envelope: {
          subject: "Simple plain text",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<simple@example.com>",
        },
        flags: new Set(),
        // Single-part messages: type is text/plain but no `part` field
        bodyStructure: { type: "text/plain" },
      });
      mockDownloadPart("Hello with em\u2014dash and caf\u00e9");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 50);

      expect(msg.body).toBe("Hello with em\u2014dash and caf\u00e9");
      expect(msg.bodyFormat).toBe("text");
      // Should default to part "1" when root has no part field
      expect(mockDownload).toHaveBeenCalledWith("50", "1", { uid: true });
    });

    it("falls back to part 1 when no text/plain or text/html found (e.g. PGP)", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 51,
        envelope: {
          subject: "Encrypted message",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<pgp@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/encrypted",
          childNodes: [
            { type: "application/pgp-encrypted", part: "1" },
            { type: "application/octet-stream", part: "2" },
          ],
        },
      });
      mockDownloadPart("-----BEGIN PGP MESSAGE-----\nsome encrypted content\n-----END PGP MESSAGE-----");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 51);

      expect(msg.body).toContain("BEGIN PGP MESSAGE");
      expect(msg.bodyFormat).toBe("text");
    });

    it("extracts attachment metadata from bodyStructure", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 47,
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
      });
      mockDownloadPart("See attached.");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 47);

      expect(msg.attachments).toHaveLength(1);
      expect(msg.attachments[0]).toEqual({
        partNumber: "2",
        filename: "report.pdf",
        contentType: "application/pdf",
        size: 12345,
      });
      expect(msg.body).toBe("See attached.");
    });

    it("handles missing envelope fields with defaults", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 48,
        envelope: {},
        flags: undefined,
        bodyStructure: undefined,
      });

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 48);

      expect(msg.subject).toBe("");
      expect(msg.from).toBe("");
      expect(msg.to).toBe("");
      expect(msg.cc).toBe("");
      expect(msg.date).toBe("");
      expect(msg.messageId).toBe("");
      expect(msg.flags).toEqual([]);
      expect(msg.body).toBe("");
      expect(msg.bodyFormat).toBe("text");
      expect(msg.attachments).toEqual([]);
    });

    it("extracts inline non-text parts as attachments", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 49,
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
      });
      mockDownloadPart("<img src='cid:logo'>");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 49);

      expect(msg.attachments).toHaveLength(1);
      expect(msg.attachments[0]).toEqual({
        partNumber: "2",
        filename: "logo.png",
        contentType: "image/png",
        size: 5000,
      });
      // HTML-only message without preferHtml → stripped
      expect(msg.bodyFormat).toBe("html-stripped");
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
    it("moves message and returns MoveResult with new UID on success", async () => {
      mockMessageMove.mockResolvedValueOnce({ uidMap: new Map([[42, 100]]) });

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");

      expect(result).toEqual({ success: true, newUid: 100, destination: "Archive" });
      expect(mockMessageMove).toHaveBeenCalledWith("42", "Archive", { uid: true });
      expect(mockConnect).toHaveBeenCalled();
      expect(mockLogout).toHaveBeenCalled();
    });

    it("returns success without newUid when uidMap is empty", async () => {
      mockMessageMove.mockResolvedValueOnce({ uidMap: new Map() });

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");

      expect(result).toEqual({ success: true, newUid: undefined, destination: "Archive" });
    });

    it("returns failure when server returns false", async () => {
      mockMessageMove.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");
      expect(result).toEqual({ success: false, destination: "Archive" });
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

  describe("markAllRead", () => {
    it("marks all unread messages as read and returns count", async () => {
      mockSearch.mockResolvedValueOnce([10, 20, 30]);

      const service = new ImapService(baseConfig);
      const count = await service.markAllRead("INBOX");

      expect(count).toBe(3);
      expect(mockSearch).toHaveBeenCalledWith({ seen: false }, { uid: true });
      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("10,20,30", ["\\Seen"], { uid: true });
    });

    it("returns 0 when no unread messages", async () => {
      mockSearch.mockResolvedValueOnce([]);

      const service = new ImapService(baseConfig);
      const count = await service.markAllRead("INBOX");

      expect(count).toBe(0);
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
    });

    it("applies olderThan filter", async () => {
      mockSearch.mockResolvedValueOnce([5]);

      const service = new ImapService(baseConfig);
      await service.markAllRead("INBOX", "2026-04-01");

      expect(mockSearch).toHaveBeenCalledWith({ seen: false, before: "2026-04-01" }, { uid: true });
    });
  });

  describe("findByMessageId", () => {
    it("returns UID when message is found", async () => {
      mockSearch.mockResolvedValueOnce([42]);

      const service = new ImapService(baseConfig);
      const uid = await service.findByMessageId("Sent", "<abc@example.com>");

      expect(uid).toBe(42);
      expect(mockSearch).toHaveBeenCalledWith({ header: { "Message-ID": "<abc@example.com>" } }, { uid: true });
    });

    it("returns undefined when not found", async () => {
      mockSearch.mockResolvedValueOnce([]);

      const service = new ImapService(baseConfig);
      const uid = await service.findByMessageId("Sent", "<missing@example.com>");

      expect(uid).toBeUndefined();
    });
  });

  describe("getThread", () => {
    it("finds thread messages by walking References headers", async () => {
      // Seed message fetch
      mockFetchOne.mockResolvedValueOnce({
        uid: 10,
        envelope: {
          messageId: "<msg10@example.com>",
          inReplyTo: "<msg5@example.com>",
          subject: "Re: Hello",
          from: [{ address: "bob@example.com" }],
          to: [{ address: "alice@example.com" }],
          date: new Date("2026-04-15T12:00:00Z"),
        },
        flags: new Set(),
        headers: Buffer.from("References: <msg5@example.com>\r\n"),
      });

      // Search for Message-ID <msg10@example.com>
      mockSearch.mockResolvedValueOnce([10]);
      // Search for References containing <msg10@example.com>
      mockSearch.mockResolvedValueOnce([]);
      // Search for In-Reply-To <msg10@example.com>
      mockSearch.mockResolvedValueOnce([]);
      // Search for Message-ID <msg5@example.com>
      mockSearch.mockResolvedValueOnce([5]);
      // Search for References containing <msg5@example.com>
      mockSearch.mockResolvedValueOnce([10]);
      // Search for In-Reply-To <msg5@example.com>
      mockSearch.mockResolvedValueOnce([10]);

      // Fetch envelopes for UIDs 5, 10
      mockFetch.mockReturnValueOnce(
        (async function* () {
          yield {
            uid: 5,
            envelope: {
              subject: "Hello",
              from: [{ address: "alice@example.com" }],
              to: [{ address: "bob@example.com" }],
              date: new Date("2026-04-14T10:00:00Z"),
            },
            flags: new Set(["\\Seen"]),
          };
          yield {
            uid: 10,
            envelope: {
              subject: "Re: Hello",
              from: [{ address: "bob@example.com" }],
              to: [{ address: "alice@example.com" }],
              date: new Date("2026-04-15T12:00:00Z"),
            },
            flags: new Set(),
          };
        })(),
      );

      const service = new ImapService(baseConfig);
      const thread = await service.getThread("INBOX", 10, 25);

      expect(thread).toHaveLength(2);
      // Sorted chronologically (oldest first)
      expect(thread[0].uid).toBe(5);
      expect(thread[0].subject).toBe("Hello");
      expect(thread[1].uid).toBe(10);
      expect(thread[1].subject).toBe("Re: Hello");
    });

    it("throws when seed message not found", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.getThread("INBOX", 999, 25)).rejects.toThrow("not found");
    });
  });

  describe("saveDraft", () => {
    it("saves draft and returns UID", async () => {
      mockAppend.mockResolvedValueOnce({ destination: "Drafts", uid: 42 });

      const service = new ImapService(baseConfig);
      const result = await service.saveDraft("Drafts", Buffer.from("Subject: Test\r\n\r\nBody"));

      expect(result).toEqual({ uid: 42 });
      expect(mockAppend).toHaveBeenCalledWith("Drafts", Buffer.from("Subject: Test\r\n\r\nBody"), [
        "\\Draft",
        "\\Seen",
      ]);
    });

    it("throws when append fails", async () => {
      mockAppend.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.saveDraft("Drafts", Buffer.from("test"))).rejects.toThrow("Failed to save draft");
    });
  });
});
