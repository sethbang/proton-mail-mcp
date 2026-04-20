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

    it("throws 'Folder not found' when getMailboxLock reports mailboxMissing", async () => {
      // imapflow sets err.mailboxMissing = true on SELECT-NO + LIST-empty (imap-flow.js:3580).
      const err = Object.assign(new Error("SELECT failed"), { mailboxMissing: true });
      mockGetMailboxLock.mockRejectedValueOnce(err);

      const service = new ImapService(baseConfig);
      await expect(service.listMessages("NoSuchFolder", 10)).rejects.toThrow("Folder not found: NoSuchFolder");

      mockGetMailboxLock.mockResolvedValue({ release: vi.fn() });
    });

    it("throws 'Folder not found' when serverResponseCode is NONEXISTENT", async () => {
      const err = Object.assign(new Error("SELECT failed"), { serverResponseCode: "NONEXISTENT" });
      mockGetMailboxLock.mockRejectedValueOnce(err);

      const service = new ImapService(baseConfig);
      await expect(service.listMessages("NoSuchFolder", 10)).rejects.toThrow("Folder not found: NoSuchFolder");

      mockGetMailboxLock.mockResolvedValue({ release: vi.fn() });
    });

    it("falls back to text matching when imapflow doesn't annotate the error", async () => {
      // Defense in depth — if imapflow's error API ever drifts, our text fallback still catches it.
      mockGetMailboxLock.mockRejectedValueOnce(new Error("No such mailbox NoSuchFolder"));

      const service = new ImapService(baseConfig);
      await expect(service.listMessages("NoSuchFolder", 10)).rejects.toThrow("Folder not found: NoSuchFolder");

      mockGetMailboxLock.mockResolvedValue({ release: vi.fn() });
    });

    it("passes through unrelated lock errors unchanged", async () => {
      mockGetMailboxLock.mockRejectedValueOnce(new Error("network blew up"));

      const service = new ImapService(baseConfig);
      await expect(service.listMessages("INBOX", 10)).rejects.toThrow("network blew up");

      mockGetMailboxLock.mockResolvedValue({ release: vi.fn() });
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

    // Regression test for v0.4.0 bug: `return this.fetchSortAndLimit(...)` without
    // await caused the outer try/finally to run client.logout() BEFORE the fetch
    // iterator resolved, closing the connection mid-fetch. Real IMAP connections
    // then threw "Connection not available" but tests passed because mocks had no
    // real connection state. This test simulates that connection state.
    it("awaits fetch to complete before closing connection", async () => {
      let logoutCalled = false;
      mockSearch.mockResolvedValueOnce([1, 2, 3]);

      mockFetch.mockReturnValueOnce(
        (async function* () {
          // Yield to the event loop so any pending finally blocks can run.
          // If the outer try/finally runs client.logout() before we resume,
          // we simulate the real imapflow "Connection not available" error.
          await new Promise((resolve) => setImmediate(resolve));
          if (logoutCalled) {
            throw new Error("Connection not available");
          }
          yield {
            uid: 1,
            envelope: { date: new Date("2026-04-15T10:00:00Z"), from: [{ address: "a@x.com" }] },
            flags: new Set(),
          };
          yield {
            uid: 2,
            envelope: { date: new Date("2026-04-14T10:00:00Z"), from: [{ address: "b@x.com" }] },
            flags: new Set(),
          };
          yield {
            uid: 3,
            envelope: { date: new Date("2026-04-13T10:00:00Z"), from: [{ address: "c@x.com" }] },
            flags: new Set(),
          };
        })(),
      );

      mockLogout.mockImplementationOnce(async () => {
        logoutCalled = true;
      });

      const service = new ImapService(baseConfig);
      const result = await service.listMessages("INBOX", 3);

      expect(result).toHaveLength(3);
      expect(logoutCalled).toBe(true); // Logout ran, but AFTER fetch completed
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

    // Same regression as listMessages — searchMessages also called fetchSortAndLimit
    // without await, causing logout to close the connection mid-fetch.
    it("awaits fetch to complete before closing connection", async () => {
      let logoutCalled = false;
      mockSearch.mockResolvedValueOnce([1, 2]);

      mockFetch.mockReturnValueOnce(
        (async function* () {
          await new Promise((resolve) => setImmediate(resolve));
          if (logoutCalled) {
            throw new Error("Connection not available");
          }
          yield {
            uid: 1,
            envelope: { date: new Date("2026-04-15T10:00:00Z"), from: [{ address: "a@x.com" }] },
            flags: new Set(),
          };
          yield {
            uid: 2,
            envelope: { date: new Date("2026-04-14T10:00:00Z"), from: [{ address: "b@x.com" }] },
            flags: new Set(),
          };
        })(),
      );

      mockLogout.mockImplementationOnce(async () => {
        logoutCalled = true;
      });

      const service = new ImapService(baseConfig);
      const result = await service.searchMessages("INBOX", { from: "test" }, 10);

      expect(result).toHaveLength(2);
      expect(logoutCalled).toBe(true);
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

    it("populates extraHeaders when showHeaders=true", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 60,
        envelope: {
          subject: "With Headers",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<hdrs@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/plain" },
        headers: Buffer.from(
          [
            "In-Reply-To: <prev@example.com>",
            "References: <first@example.com>",
            "Reply-To: reply@example.com",
            "List-Unsubscribe: <mailto:unsub@example.com>",
            "",
            "",
          ].join("\r\n"),
        ),
      });
      mockDownloadPart("Body");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 60, { showHeaders: true });

      expect(msg.extraHeaders).toBeDefined();
      expect(msg.extraHeaders?.["in-reply-to"]).toBe("<prev@example.com>");
      expect(msg.extraHeaders?.["references"]).toBe("<first@example.com>");
      expect(msg.extraHeaders?.["reply-to"]).toBe("reply@example.com");
      expect(msg.extraHeaders?.["list-unsubscribe"]).toBe("<mailto:unsub@example.com>");

      // The fetch must have requested these headers
      const fetchOptions = mockFetchOne.mock.calls[0][1];
      expect(fetchOptions).toHaveProperty("headers");
      const requested = (fetchOptions.headers as string[]).map((h: string) => h.toLowerCase());
      expect(requested).toContain("in-reply-to");
      expect(requested).toContain("references");
      expect(requested).toContain("reply-to");
      expect(requested).toContain("list-unsubscribe");
    });

    it("unfolds continuation lines in header values", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 61,
        envelope: {
          subject: "Folded",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<folded@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/plain" },
        headers: Buffer.from(
          [
            "References: <a@example.com>",
            "\t<b@example.com>",
            " <c@example.com>",
            "In-Reply-To: <a@example.com>",
            "",
            "",
          ].join("\r\n"),
        ),
      });
      mockDownloadPart("Body");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 61, { showHeaders: true });

      expect(msg.extraHeaders?.["references"]).toContain("<a@example.com>");
      expect(msg.extraHeaders?.["references"]).toContain("<b@example.com>");
      expect(msg.extraHeaders?.["references"]).toContain("<c@example.com>");
    });

    it("does not request extra headers when showHeaders is not set", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 62,
        envelope: {
          subject: "Default",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
          messageId: "<default@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/plain" },
      });
      mockDownloadPart("Body");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 62);

      expect(msg.extraHeaders).toBeUndefined();
      const fetchOptions = mockFetchOne.mock.calls[0][1];
      expect(fetchOptions.headers).toBeFalsy();
    });

    // Round 3: the round-2 findPartNumber walked by content-type only, so a text/plain
    // *attachment* (disposition=attachment) sitting next to an HTML body was picked as
    // the body — silently returning attachment content to the caller.
    it("picks the HTML body, not a sibling text/plain attachment", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 70,
        envelope: {
          subject: "HTML + text attachment",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-19T12:00:00Z"),
          messageId: "<htmlattach@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            {
              type: "multipart/alternative",
              childNodes: [
                { type: "text/plain", part: "1.1" },
                { type: "text/html", part: "1.2" },
              ],
            },
            {
              type: "text/plain",
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "hello.txt" },
              size: 31,
            },
          ],
        },
      });
      // Expect the body to come from part 1.1 (text/plain inside multipart/alternative),
      // NOT part 2 (the attachment).
      mockDownloadPart("This is the actual body text.");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 70);

      expect(msg.body).toBe("This is the actual body text.");
      expect(mockDownload).toHaveBeenCalledWith("70", "1.1", { uid: true });
      // Regression guard: the attachment still appears in attachments
      expect(msg.attachments).toEqual([
        { partNumber: "2", filename: "hello.txt", contentType: "text/plain", size: 31 },
      ]);
    });

    it("strips HTML when the only body is HTML and a text/plain attachment sits alongside", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 71,
        envelope: {
          subject: "HTML-only + attachment",
          from: [{ address: "sender@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-19T12:00:00Z"),
          messageId: "<htmlonlyattach@example.com>",
        },
        flags: new Set(),
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/html", part: "1" },
            {
              type: "text/plain",
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "hello.txt" },
              size: 31,
            },
          ],
        },
      });
      mockDownloadPart("<html><body><p>Body HTML content</p></body></html>");

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 71);

      // Should strip HTML from part 1, not read the text/plain attachment (part 2).
      expect(msg.bodyFormat).toBe("html-stripped");
      expect(msg.body).toContain("Body HTML content");
      expect(mockDownload).toHaveBeenCalledWith("71", "1", { uid: true });
    });

    it("stripUrls=true drops anchor href when rendering HTML-only body", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 72,
        envelope: {
          subject: "Newsletter",
          from: [{ address: "newsletter@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-19T12:00:00Z"),
          messageId: "<news@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/html", part: "1" },
      });
      mockDownloadPart(
        '<html><body><p>Check out <a href="https://tracker.example/c/verylongpath?id=abc123">our new thing</a>!</p></body></html>',
      );

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 72, { stripUrls: true });

      expect(msg.body).toContain("our new thing");
      expect(msg.body).not.toContain("tracker.example");
      expect(msg.body).not.toContain("https://");
    });

    it("stripUrls default preserves href in brackets (regression guard)", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 73,
        envelope: {
          subject: "Newsletter",
          from: [{ address: "newsletter@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-19T12:00:00Z"),
          messageId: "<news2@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/html", part: "1" },
      });
      mockDownloadPart(
        '<html><body><p>Check out <a href="https://tracker.example/c/verylongpath">our new thing</a>!</p></body></html>',
      );

      const service = new ImapService(baseConfig);
      const msg = await service.readMessage("INBOX", 73);

      expect(msg.body).toContain("our new thing");
      expect(msg.body).toContain("https://tracker.example/c/verylongpath");
    });
  });

  describe("listAttachments", () => {
    it("returns attachment metadata without downloading the body", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
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
            {
              type: "image/png",
              part: "3",
              disposition: "attachment",
              dispositionParameters: { filename: "chart.png" },
              size: 6789,
            },
          ],
        },
      });

      const service = new ImapService(baseConfig);
      const atts = await service.listAttachments("INBOX", 42);

      expect(atts).toEqual([
        { partNumber: "2", filename: "report.pdf", contentType: "application/pdf", size: 12345 },
        { partNumber: "3", filename: "chart.png", contentType: "image/png", size: 6789 },
      ]);
      // Body must not be downloaded
      expect(mockDownload).not.toHaveBeenCalled();
    });

    it("returns empty array when the message has no attachments", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        bodyStructure: { type: "text/plain", part: "1" },
      });

      const service = new ImapService(baseConfig);
      const atts = await service.listAttachments("INBOX", 42);
      expect(atts).toEqual([]);
    });

    it("throws when UID does not exist", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.listAttachments("INBOX", 999999)).rejects.toThrow("Message UID 999999 not found in INBOX");
    });

    it("rejects folder path with control characters", async () => {
      const service = new ImapService(baseConfig);
      await expect(service.listAttachments("IN\r\nBOX", 1)).rejects.toThrow("invalid control characters");
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
    // Helper to mock the pre-check fetch of bodyStructure so downloadAttachment
    // can validate the partNumber before calling client.download.
    function mockBodyStructure(parts: { part: string; filename?: string; type?: string }[]) {
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/plain", part: "0" },
            ...parts.map((p) => ({
              type: p.type ?? "application/octet-stream",
              part: p.part,
              disposition: "attachment",
              dispositionParameters: { filename: p.filename ?? `file-${p.part}` },
              size: 1,
            })),
          ],
        },
      });
    }

    it("downloads and returns base64-encoded content", async () => {
      mockBodyStructure([{ part: "2", filename: "report.pdf", type: "application/pdf" }]);
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
      mockBodyStructure([{ part: "1", filename: "big.txt", type: "text/plain" }]);
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
      mockBodyStructure([{ part: "3" }]);
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

    it("throws a clear error when the part number is not on the message", async () => {
      // bodyStructure only has part 2; user asks for 42
      mockBodyStructure([{ part: "2", filename: "real.pdf", type: "application/pdf" }]);

      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("INBOX", 42, "42")).rejects.toThrow(
        /Part 42 not found on UID 42 in INBOX.*known parts.*2/,
      );
      expect(mockDownload).not.toHaveBeenCalled();
    });

    it("throws when UID does not exist", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("INBOX", 999999, "1")).rejects.toThrow(
        "Message UID 999999 not found in INBOX",
      );
      expect(mockDownload).not.toHaveBeenCalled();
    });

    it("reports 'known parts: [(none)]' when the message has no attachments at all", async () => {
      // Message is a plain body-only email with no attachments. Any partNumber lookup
      // must surface the empty list so the caller knows there's nothing to download.
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        bodyStructure: { type: "text/plain", part: "1" },
      });

      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("INBOX", 42, "1")).rejects.toThrow(
        "Part 1 not found on UID 42 in INBOX; known parts: [(none)]",
      );
      expect(mockDownload).not.toHaveBeenCalled();
    });

    it("lists all attachment parts in the error when the requested part is missing", async () => {
      mockBodyStructure([
        { part: "2", filename: "a.pdf", type: "application/pdf" },
        { part: "3", filename: "b.png", type: "image/png" },
        { part: "4", filename: "c.txt", type: "text/plain" },
      ]);

      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("INBOX", 42, "99")).rejects.toThrow(
        "Part 99 not found on UID 42 in INBOX; known parts: [2, 3, 4]",
      );
    });
  });

  // Regression guard for round-3 "forward_email sets \Answered on source" — nothing in
  // the forward path (readMessage → downloadAttachment → sendEmail) should ever touch flags.
  // If \Answered appears on the source after a forward, it's Proton Bridge behavior, not us.
  describe("flag-setting side-effects (regression guard)", () => {
    it("readMessage never calls messageFlagsAdd / messageFlagsRemove", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        envelope: {
          subject: "Test",
          from: [{ address: "a@example.com" }],
          to: [{ address: "b@example.com" }],
          date: new Date("2026-04-20T12:00:00Z"),
          messageId: "<test@example.com>",
        },
        flags: new Set(),
        bodyStructure: { type: "text/plain", part: "1" },
      });
      mockDownload.mockResolvedValueOnce({
        meta: {},
        content: (async function* () {
          yield Buffer.from("body");
        })(),
      });

      const service = new ImapService(baseConfig);
      await service.readMessage("INBOX", 42);

      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
      expect(mockMessageFlagsRemove).not.toHaveBeenCalled();
    });

    it("downloadAttachment never calls messageFlagsAdd / messageFlagsRemove", async () => {
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/plain", part: "1" },
            {
              type: "application/pdf",
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "a.pdf" },
              size: 10,
            },
          ],
        },
      });
      mockDownload.mockResolvedValueOnce({
        meta: { contentType: "application/pdf", filename: "a.pdf" },
        content: (async function* () {
          yield Buffer.from("pdf");
        })(),
      });

      const service = new ImapService(baseConfig);
      await service.downloadAttachment("INBOX", 42, "2");

      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
      expect(mockMessageFlagsRemove).not.toHaveBeenCalled();
    });
  });

  describe("moveMessage", () => {
    it("moves message and returns MoveResult with new UID on success", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      mockMessageMove.mockResolvedValueOnce({ uidMap: new Map([[42, 100]]) });

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");

      expect(result).toEqual({ success: true, newUid: 100, destination: "Archive" });
      expect(mockMessageMove).toHaveBeenCalledWith("42", "Archive", { uid: true });
      expect(mockConnect).toHaveBeenCalled();
      expect(mockLogout).toHaveBeenCalled();
    });

    it("returns success without newUid when uidMap is empty", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      mockMessageMove.mockResolvedValueOnce({ uidMap: new Map() });

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");

      expect(result).toEqual({ success: true, newUid: undefined, destination: "Archive" });
    });

    // Superseded by the destination-probe tests below: when messageMove returns false
    // we now probe the destination to distinguish "missing folder" from "other failure".

    it("throws when source UID does not exist in the folder", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.moveMessage("INBOX", 999999, "Archive")).rejects.toThrow(
        "Message UID 999999 not found in INBOX",
      );
      expect(mockMessageMove).not.toHaveBeenCalled();
    });

    // imapflow's messageMove swallows COPY failures internally (copy.js:42-46) and returns
    // `false` — the underlying TRYCREATE/NONEXISTENT error never surfaces. So a try/catch
    // around messageMove can never fire; we probe the destination with status() on failure.
    it("probes destination with status() when messageMove returns false and translates missing folder", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      mockMessageMove.mockResolvedValueOnce(false);
      const err = Object.assign(new Error("STATUS failed"), { mailboxMissing: true });
      mockStatus.mockRejectedValueOnce(err);

      const service = new ImapService(baseConfig);
      await expect(service.moveMessage("INBOX", 42, "NonexistentFolder")).rejects.toThrow(
        "Destination folder not found: NonexistentFolder",
      );
      expect(mockStatus).toHaveBeenCalledWith("NonexistentFolder", { messages: true });
    });

    it("keeps the generic failure when destination exists but move still fails", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      mockMessageMove.mockResolvedValueOnce(false);
      mockStatus.mockResolvedValueOnce({ messages: 5 }); // probe succeeds — folder exists

      const service = new ImapService(baseConfig);
      const result = await service.moveMessage("INBOX", 42, "Archive");
      expect(result).toEqual({ success: false, destination: "Archive" });
    });

    it("translates source-folder-missing via mailboxMissing signal", async () => {
      const err = Object.assign(new Error("SELECT failed"), { mailboxMissing: true });
      mockGetMailboxLock.mockRejectedValueOnce(err);

      const service = new ImapService(baseConfig);
      await expect(service.moveMessage("NoSuchFolder", 42, "Archive")).rejects.toThrow(
        "Folder not found: NoSuchFolder",
      );

      mockGetMailboxLock.mockResolvedValue({ release: vi.fn() });
    });
  });

  describe("deleteMessage", () => {
    it("deletes message and returns true on success", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });

      const service = new ImapService(baseConfig);
      const result = await service.deleteMessage("INBOX", 42);

      expect(result).toBe(true);
      expect(mockMessageDelete).toHaveBeenCalledWith("42", { uid: true });
      expect(mockConnect).toHaveBeenCalled();
      expect(mockLogout).toHaveBeenCalled();
    });

    it("returns false when server returns false", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      mockMessageDelete.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      const result = await service.deleteMessage("INBOX", 42);
      expect(result).toBe(false);
    });

    it("throws when UID does not exist and does not call messageDelete", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.deleteMessage("INBOX", 999999)).rejects.toThrow("Message UID 999999 not found in INBOX");
      expect(mockMessageDelete).not.toHaveBeenCalled();
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
      // Pre-check: part 2 must exist on the bodyStructure.
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/plain", part: "1" },
            {
              type: "application/octet-stream",
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "huge.bin" },
              size: 30 * 1024 * 1024,
            },
          ],
        },
      });

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
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      const service = new ImapService(baseConfig);
      await service.updateFlags("INBOX", 42, ["\\Seen", "\\Flagged"], []);

      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("42", ["\\Seen", "\\Flagged"], { uid: true });
      expect(mockMessageFlagsRemove).not.toHaveBeenCalled();
    });

    it("removes flags when flagsToRemove is non-empty", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      const service = new ImapService(baseConfig);
      await service.updateFlags("INBOX", 42, [], ["\\Seen"]);

      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("42", ["\\Seen"], { uid: true });
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
    });

    it("adds and removes flags in one call", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 });
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 42, ["\\Flagged"], ["\\Seen"]);

      expect(result).toBe(true);
      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("42", ["\\Flagged"], { uid: true });
      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("42", ["\\Seen"], { uid: true });
    });

    it("throws when UID does not exist and does not call STORE", async () => {
      mockFetchOne.mockResolvedValueOnce(false);

      const service = new ImapService(baseConfig);
      await expect(service.updateFlags("INBOX", 999999, ["\\Seen"], [])).rejects.toThrow(
        "Message UID 999999 not found in INBOX",
      );
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
      expect(mockMessageFlagsRemove).not.toHaveBeenCalled();
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

    it("resolves seed by messageId across default folders when messageId is supplied", async () => {
      // When messageId is supplied, the service should search for it across the default folder
      // set (INBOX, Sent, All Mail) instead of trusting a (uid, folder) pair.
      //
      // Search call sequence (in order):
      //  1. findByMessageId in INBOX → empty
      //  2. findByMessageId in Sent → returns [7]
      //  (no need to check All Mail once we've located the seed)
      //  Then seed fetchOne in the locating folder (Sent),
      //  then the reference walk across folders.
      mockSearch
        .mockResolvedValueOnce([]) // INBOX: no match
        .mockResolvedValueOnce([7]) // Sent: match
        // Reference walk: Message-ID in INBOX, Sent, All Mail; then References; then In-Reply-To
        .mockResolvedValue([]);

      mockFetchOne.mockResolvedValueOnce({
        uid: 7,
        envelope: {
          messageId: "<seed@example.com>",
          inReplyTo: "",
          subject: "Re: hi",
          from: [{ address: "me@pm.me" }],
          to: [{ address: "you@example.com" }],
          date: new Date("2026-04-15T12:00:00Z"),
        },
        flags: new Set(),
        headers: Buffer.from(""),
      });

      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 7,
            envelope: {
              subject: "Re: hi",
              from: [{ address: "me@pm.me" }],
              to: [{ address: "you@example.com" }],
              date: new Date("2026-04-15T12:00:00Z"),
            },
            flags: new Set(),
          };
        })(),
      );

      const service = new ImapService(baseConfig);
      const thread = await service.getThreadByMessageId("<seed@example.com>", 25);

      expect(thread.length).toBeGreaterThan(0);
      expect(thread[0].uid).toBe(7);
      // Should have locked INBOX and Sent at minimum while searching
      const lockedFolders = mockGetMailboxLock.mock.calls.map((c) => c[0]);
      expect(lockedFolders).toContain("INBOX");
      expect(lockedFolders).toContain("Sent");
    });

    it("throws when messageId is not found in any default folder", async () => {
      // All default folders return empty for the Message-ID search
      mockSearch.mockResolvedValue([]);

      const service = new ImapService(baseConfig);
      await expect(service.getThreadByMessageId("<missing@example.com>", 25)).rejects.toThrow("not found");
    });

    it("skips folders that don't exist while walking default set", async () => {
      // All Mail doesn't exist on every provider — a NONEXISTENT there must not kill
      // the whole thread lookup. Simulate by having getMailboxLock reject for "All Mail".
      const originalLock = mockGetMailboxLock.getMockImplementation();
      mockGetMailboxLock.mockImplementation((folder: string) => {
        if (folder === "All Mail") {
          const err = Object.assign(new Error("SELECT failed"), { mailboxMissing: true });
          return Promise.reject(err);
        }
        return Promise.resolve({ release: vi.fn() });
      });

      mockSearch.mockResolvedValueOnce([]).mockResolvedValueOnce([3]).mockResolvedValue([]);

      mockFetchOne.mockResolvedValueOnce({
        uid: 3,
        envelope: {
          messageId: "<seed2@example.com>",
          inReplyTo: "",
          subject: "hi",
          from: [{ address: "me@pm.me" }],
          to: [{ address: "you@example.com" }],
          date: new Date("2026-04-15T12:00:00Z"),
        },
        flags: new Set(),
        headers: Buffer.from(""),
      });

      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 3,
            envelope: {
              subject: "hi",
              from: [{ address: "me@pm.me" }],
              to: [{ address: "you@example.com" }],
              date: new Date("2026-04-15T12:00:00Z"),
            },
            flags: new Set(),
          };
        })(),
      );

      const service = new ImapService(baseConfig);
      // Should not throw despite All Mail being missing
      const thread = await service.getThreadByMessageId("<seed2@example.com>", 25);
      expect(thread.length).toBeGreaterThan(0);

      // Restore implementation
      mockGetMailboxLock.mockImplementation(originalLock ?? (() => Promise.resolve({ release: vi.fn() })));
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

    it("translates mailboxMissing on append to 'Folder not found'", async () => {
      const err = Object.assign(new Error("APPEND failed"), { mailboxMissing: true });
      mockAppend.mockRejectedValueOnce(err);

      const service = new ImapService(baseConfig);
      await expect(service.saveDraft("NoSuchDrafts", Buffer.from("test"))).rejects.toThrow(
        "Folder not found: NoSuchDrafts",
      );
    });
  });

  // Every folder-taking service method must translate imapflow's mailboxMissing / NONEXISTENT /
  // TRYCREATE signals into a uniform "Folder not found: <path>" error. Round-1 fixed listMessages
  // via text matching (which didn't actually catch the real imapflow errors); round-2 routes
  // every call through a shared lockFolder() helper and this block pins that contract.
  describe("folder-not-found translation for all folder-taking tools", () => {
    function rejectMailboxMissingOnce() {
      const err = Object.assign(new Error("SELECT failed"), { mailboxMissing: true });
      mockGetMailboxLock.mockRejectedValueOnce(err);
    }

    function resetLockDefault() {
      mockGetMailboxLock.mockResolvedValue({ release: vi.fn() });
    }

    it("searchMessages", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.searchMessages("NoSuchFolder", { from: "a" }, 10)).rejects.toThrow(
        "Folder not found: NoSuchFolder",
      );
      resetLockDefault();
    });

    it("readMessage", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.readMessage("NoSuchFolder", 1)).rejects.toThrow("Folder not found: NoSuchFolder");
      resetLockDefault();
    });

    it("downloadAttachment", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.downloadAttachment("NoSuchFolder", 1, "1")).rejects.toThrow(
        "Folder not found: NoSuchFolder",
      );
      resetLockDefault();
    });

    it("deleteMessage", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.deleteMessage("NoSuchFolder", 1)).rejects.toThrow("Folder not found: NoSuchFolder");
      resetLockDefault();
    });

    it("updateFlags", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.updateFlags("NoSuchFolder", 1, ["\\Seen"], [])).rejects.toThrow(
        "Folder not found: NoSuchFolder",
      );
      resetLockDefault();
    });

    it("markAllRead", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.markAllRead("NoSuchFolder")).rejects.toThrow("Folder not found: NoSuchFolder");
      resetLockDefault();
    });

    it("getThread (uid path)", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.getThread("NoSuchFolder", 1, 25)).rejects.toThrow("Folder not found: NoSuchFolder");
      resetLockDefault();
    });

    it("findByMessageId", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.findByMessageId("NoSuchFolder", "<a@b>")).rejects.toThrow("Folder not found: NoSuchFolder");
      resetLockDefault();
    });

    it("listAttachments", async () => {
      rejectMailboxMissingOnce();
      const service = new ImapService(baseConfig);
      await expect(service.listAttachments("NoSuchFolder", 1)).rejects.toThrow("Folder not found: NoSuchFolder");
      resetLockDefault();
    });
  });
});
