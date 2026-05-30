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
const mockMailboxCreate = vi.fn();
const mockMailboxRename = vi.fn();
const mockMailboxDelete = vi.fn();
const mockMessageCopy = vi.fn();

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
      messageCopy: mockMessageCopy,
      download: mockDownload,
      append: mockAppend,
      mailboxCreate: mockMailboxCreate,
      mailboxRename: mockMailboxRename,
      mailboxDelete: mockMailboxDelete,
    };
  }),
}));

const baseConfig: ImapConfig = {
  host: "127.0.0.1",
  port: 1143,
  secure: false,
  auth: { user: "test@pm.me", pass: "test-bridge-password" },
};

function makeService() {
  return new ImapService(baseConfig);
}

/**
 * Yield an async iterator of `{ uid }` objects, used to mock the FETCH-based
 * existence pre-check (`existingUidSet`). FETCH only yields existing UIDs, so the
 * passed array represents which UIDs the mailbox actually contains.
 *
 * Use with `mockFetch.mockImplementation(() => uidIter([...]))` so each call gets
 * a fresh, un-exhausted iterator — `mockReturnValue` reuses the same iterator and
 * the second call sees an empty stream.
 */
function uidIter(uids: number[]): AsyncGenerator<{ uid: number }, void, unknown> {
  return (async function* () {
    for (const u of uids) yield { uid: u };
  })();
}

/**
 * Yield an async iterator of `{ uid, flags: Set<string> }`. Used to mock the
 * post-STORE verify FETCH for bulkUpdateFlags — each yielded object carries
 * the flag state the test wants the server to claim AFTER the STORE.
 */
function uidIterWithFlags(
  entries: Array<[number, string[]]>,
): AsyncGenerator<{ uid: number; flags: Set<string> }, void, unknown> {
  return (async function* () {
    for (const [uid, flags] of entries) yield { uid, flags: new Set(flags) };
  })();
}

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
        noSelect: false,
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
        noSelect: false,
      });
      expect(folders[1]).toEqual({
        path: "Drafts",
        name: "Drafts",
        specialUse: "",
        messages: 0,
        unseen: 0,
        noSelect: false,
      });
    });

    it("handles empty mailbox list", async () => {
      mockList.mockResolvedValueOnce([]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders).toEqual([]);
    });

    it("tags Noselect namespace containers", async () => {
      // Proton exposes "Folders" and "Labels" as top-level namespace
      // containers — they hold nested children but can't be opened or used
      // as move destinations. imapflow surfaces this via `flags` containing
      // `\Noselect` or `\NonExistent`, depending on bridge version.
      mockList.mockResolvedValueOnce([
        {
          path: "Folders",
          name: "Folders",
          specialUse: "",
          flags: new Set(["\\Noselect", "\\HasChildren"]),
          status: { messages: 0, unseen: 0 },
        },
        {
          path: "Labels",
          name: "Labels",
          specialUse: "",
          flags: new Set(["\\NonExistent", "\\HasChildren"]),
          status: { messages: 0, unseen: 0 },
        },
        { path: "INBOX", name: "Inbox", specialUse: "\\Inbox", status: { messages: 5, unseen: 1 } },
      ]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders[0].noSelect).toBe(true);
      expect(folders[1].noSelect).toBe(true);
      expect(folders[2].noSelect).toBe(false);
    });

    it("does NOT tag a selectable label as a namespace just because listed=false", async () => {
      // Proton surfaces labels via LSUB with `listed: false`, yet they are fully
      // selectable — they hold messages, accept COPY, and report live counts.
      // imapflow proves this by attaching a `status` object (it only does so for
      // non-\Noselect mailboxes). The old `listed === false` heuristic tagged
      // this populated label "(namespace — not a mailbox)", contradicting the
      // message count on the same line. It must come back as selectable.
      mockList.mockResolvedValueOnce([
        {
          path: "Labels/Receipts",
          name: "Receipts",
          specialUse: "",
          listed: false,
          status: { messages: 3, unseen: 1 },
        },
        // A genuine container that lacks the \Noselect flag but also has no
        // status (imapflow couldn't STATUS it) is still caught by the narrow
        // gated branch.
        { path: "Folders", name: "Folders", specialUse: "", listed: false, status: undefined },
      ]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders[0].noSelect).toBe(false);
      expect(folders[0].messages).toBe(3);
      expect(folders[1].noSelect).toBe(true);
    });

    it("does NOT tag a populated label as a namespace even when the bridge sets a \\Noselect/\\NonExistent flag", async () => {
      // Proton Mail Bridge has been observed returning a populated label with
      // BOTH a `\Noselect`-family flag AND inline STATUS counts (LIST-STATUS).
      // A live message count proves selectability, so it must override the flag —
      // otherwise `Labels/X — 3 messages` renders as "(namespace — not a mailbox)",
      // contradicting the count on the same line. Reconstructed from the observed
      // wire state. Both flag variants are covered since the bridge may use either.
      mockList.mockResolvedValueOnce([
        {
          path: "Labels/Newsletters",
          name: "Newsletters",
          specialUse: "",
          flags: new Set(["\\Noselect", "\\HasNoChildren"]),
          status: { messages: 3, unseen: 1 },
        },
        {
          path: "Labels/Receipts",
          name: "Receipts",
          specialUse: "",
          flags: new Set(["\\NonExistent", "\\HasNoChildren"]),
          status: { messages: 3, unseen: 1 },
        },
      ]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders[0].noSelect).toBe(false);
      expect(folders[0].messages).toBe(3);
      expect(folders[1].noSelect).toBe(false);
      expect(folders[1].messages).toBe(3);
    });

    it("does NOT tag an EMPTY leaf label as a namespace even when mis-flagged \\Noselect", async () => {
      // The messages>0 override can't rescue an empty label the bridge wrongly
      // flags \Noselect. A successful STATUS on a childless (\HasNoChildren)
      // mailbox is the additional selectability signal — a STATUS only succeeds
      // on an openable mailbox, and a genuine container carries \HasChildren.
      mockList.mockResolvedValueOnce([
        // Empty leaf label, mis-flagged \Noselect — STATUS succeeded (status present, 0 messages).
        {
          path: "Labels/Empty",
          name: "Empty",
          specialUse: "",
          flags: new Set(["\\Noselect", "\\HasNoChildren"]),
          status: { messages: 0, unseen: 0 },
        },
        // Genuine empty container — has children, must stay tagged.
        {
          path: "Folders",
          name: "Folders",
          specialUse: "",
          flags: new Set(["\\Noselect", "\\HasChildren"]),
          status: { messages: 0, unseen: 0 },
        },
      ]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders[0].noSelect).toBe(false); // empty leaf label is selectable
      expect(folders[1].noSelect).toBe(true); // namespace container stays tagged
    });

    it("still tags a \\Noselect container when the server omits the CHILDREN extension flags", async () => {
      // Defense for the selectability rule: it keys off the *explicit*
      // \HasNoChildren flag, not the absence of \HasChildren. A server that
      // doesn't advertise the CHILDREN extension (RFC 3348) returns the
      // container with neither flag; the absence-of-\HasChildren form would
      // mistake it for a selectable empty leaf. Explicit-presence keeps it
      // tagged.
      mockList.mockResolvedValueOnce([
        {
          path: "Folders",
          name: "Folders",
          specialUse: "",
          flags: new Set(["\\Noselect"]), // neither \HasChildren nor \HasNoChildren
          status: { messages: 0, unseen: 0 },
        },
      ]);
      const service = new ImapService(baseConfig);
      const folders = await service.listFolders();
      expect(folders[0].noSelect).toBe(true);
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

    it("rejects a non-selectable namespace container instead of returning empty", async () => {
      // Proton lets SELECT succeed on its top-level "Folders"/"Labels" nodes and
      // reports them empty, so a zero-result read can't distinguish "empty
      // mailbox" from "not a mailbox". On the zero-result path the tool consults
      // LIST and rejects a \Noselect container with an actionable error.
      mockSearch.mockResolvedValueOnce([]); // SELECT succeeds, no messages
      mockList.mockResolvedValueOnce([
        {
          path: "Folders",
          name: "Folders",
          specialUse: "",
          flags: new Set(["\\Noselect", "\\HasChildren"]),
          status: { messages: 0, unseen: 0 },
        },
      ]);
      const service = new ImapService(baseConfig);
      await expect(service.listMessages("Folders", 10)).rejects.toThrow(
        /namespace container, not a selectable mailbox/,
      );
    });

    it("does NOT reject a genuinely empty, selectable folder on zero results", async () => {
      mockSearch.mockResolvedValueOnce([]);
      mockList.mockResolvedValueOnce([
        {
          path: "Folders/Empty",
          name: "Empty",
          specialUse: "",
          flags: new Set(["\\HasNoChildren"]),
          status: { messages: 0, unseen: 0 },
        },
      ]);
      const service = new ImapService(baseConfig);
      await expect(service.listMessages("Folders/Empty", 10)).resolves.toEqual([]);
    });

    it("sortByUid orders by UID descending regardless of message date", async () => {
      // Selects the highest `limit` UIDs (4,5) and orders by UID — NOT date.
      // UID 4 is dated AFTER UID 5, so a date sort would put 4 first; UID order must win.
      mockSearch.mockResolvedValueOnce([1, 2, 3, 4, 5]);
      mockFetch.mockReturnValueOnce(
        (async function* () {
          yield {
            uid: 4,
            envelope: { subject: "u4", from: [{ address: "a@b" }], date: new Date("2026-05-10") },
            flags: new Set(),
          };
          yield {
            uid: 5,
            envelope: { subject: "u5", from: [{ address: "a@b" }], date: new Date("2026-05-01") },
            flags: new Set(),
          };
        })(),
      );
      const service = new ImapService(baseConfig);
      const result = await service.listMessages("All Mail", 2, undefined, { sortByUid: true });
      expect(result.map((m) => m.uid)).toEqual([5, 4]);
    });

    it("sortByUid pagination covers boundary messages that date-mode pagination would skip", async () => {
      // All Mail: UID 1 is dated LATER (19:10) than UID 2 (19:05) — UID order ≠ date order.
      // Date mode would return UID 1 first (cursor beforeUid=1) and skip UID 2. sortByUid
      // pages by UID: page 1 → UID 2 (highest), page 2 (beforeUid=2) → UID 1. No skip.
      const service = new ImapService(baseConfig);

      mockSearch.mockResolvedValueOnce([1, 2]); // page 1: search(all)
      mockFetch.mockReturnValueOnce(
        (async function* () {
          yield {
            uid: 2,
            envelope: {
              subject: "hi-uid earlier-date",
              from: [{ address: "a@b" }],
              date: new Date("2026-05-29T19:05:00Z"),
            },
            flags: new Set(),
          };
        })(),
      );
      const page1 = await service.listMessages("All Mail", 1, undefined, { sortByUid: true });
      expect(page1.map((m) => m.uid)).toEqual([2]);

      mockSearch.mockResolvedValueOnce([1]); // page 2: search(uid 1:1)
      mockFetch.mockReturnValueOnce(
        (async function* () {
          yield {
            uid: 1,
            envelope: {
              subject: "lo-uid later-date",
              from: [{ address: "a@b" }],
              date: new Date("2026-05-29T19:10:00Z"),
            },
            flags: new Set(),
          };
        })(),
      );
      const page2 = await service.listMessages("All Mail", 1, 2, { sortByUid: true });
      expect(page2.map((m) => m.uid)).toEqual([1]);

      // Union covers both UIDs — neither skipped nor duplicated.
      expect([...page1, ...page2].map((m) => m.uid).sort()).toEqual([1, 2]);
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
        encodedSize: 12345,
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
        encodedSize: 5000,
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

    // Regression guard: an earlier findPartNumber walked by content-type only, so a
    // text/plain *attachment* (disposition=attachment) sitting next to an HTML body was
    // picked as the body — silently returning attachment content to the caller.
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
        { partNumber: "2", filename: "hello.txt", contentType: "text/plain", size: 31, encodedSize: 31 },
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
        { partNumber: "2", filename: "report.pdf", contentType: "application/pdf", size: 12345, encodedSize: 12345 },
        { partNumber: "3", filename: "chart.png", contentType: "image/png", size: 6789, encodedSize: 6789 },
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

    it("reports decoded file size for base64 parts, keeping the encoded size separately", async () => {
      // Reporting bug: IMAP reports the *encoded* octet count. For a
      // base64 part that's ~37% larger than the real file. `size` must be the
      // decoded estimate (what the user sees on save); `encodedSize` keeps the
      // wire size for the inline cap. 58068 encoded → 42434 decoded (a real
      // pair observed against Proton Bridge).
      mockFetchOne.mockResolvedValueOnce({
        uid: 42,
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/plain", part: "1" },
            {
              type: "text/markdown",
              part: "2",
              encoding: "base64",
              disposition: "attachment",
              dispositionParameters: { filename: "README.md" },
              size: 58068,
            },
          ],
        },
      });

      const service = new ImapService(baseConfig);
      const atts = await service.listAttachments("INBOX", 42);
      expect(atts).toEqual([
        { partNumber: "2", filename: "README.md", contentType: "text/markdown", size: 42434, encodedSize: 58068 },
      ]);
    });

    it("decodedSizeEstimate: base64 conversion (case-insensitive) and non-base64 passthrough", () => {
      // Both encoded→decoded pairs observed against Proton Bridge, exactly.
      expect(ImapService.decodedSizeEstimate(58068, "base64")).toBe(42434);
      expect(ImapService.decodedSizeEstimate(77440, "base64")).toBe(56590);
      // Encoding token casing varies by server; must still trigger the branch.
      expect(ImapService.decodedSizeEstimate(58068, "BASE64")).toBe(42434);
      // Non-base64 transfer encodings have encoded ≈ decoded → passthrough.
      expect(ImapService.decodedSizeEstimate(5000, "7bit")).toBe(5000);
      expect(ImapService.decodedSizeEstimate(5000, "quoted-printable")).toBe(5000);
      expect(ImapService.decodedSizeEstimate(5000, undefined)).toBe(5000);
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

  // Regression guard for "forward_email sets \Answered on source". Nothing in the
  // forward path (readMessage → downloadAttachment → sendEmail) STOREs flags directly —
  // this guard pins that. The root cause was indirect: the forward
  // handler used to send `In-Reply-To`/`References: <original>`, which made Proton's
  // server set \Answered on the source (In-Reply-To declares "this answers X"). The fix
  // is in index.ts (forward no longer sends those threading headers); the server-side
  // \Answered behavior itself is only confirmable against a live Bridge.
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
    it("adds flags when flagsToAdd is non-empty and verifies via post-STORE FETCH", async () => {
      mockFetchOne
        .mockResolvedValueOnce({ uid: 42 }) // existence check
        .mockResolvedValueOnce({ uid: 42, flags: new Set(["\\Seen", "\\Flagged"]) }); // post-STORE verify
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 42, ["\\Seen", "\\Flagged"], []);

      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("42", ["\\Seen", "\\Flagged"], { uid: true });
      expect(mockMessageFlagsRemove).not.toHaveBeenCalled();
      expect(result).toEqual({ added: ["\\Seen", "\\Flagged"], removed: [], notApplied: [] });
    });

    it("removes flags when flagsToRemove is non-empty and verifies via post-STORE FETCH", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 }).mockResolvedValueOnce({ uid: 42, flags: new Set() }); // verified \Seen absent
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 42, [], ["\\Seen"]);

      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("42", ["\\Seen"], { uid: true });
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
      expect(result).toEqual({ added: [], removed: ["\\Seen"], notApplied: [] });
    });

    it("adds and removes flags in one call", async () => {
      mockFetchOne.mockResolvedValueOnce({ uid: 42 }).mockResolvedValueOnce({ uid: 42, flags: new Set(["\\Flagged"]) });
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 42, ["\\Flagged"], ["\\Seen"]);

      expect(result).toEqual({ added: ["\\Flagged"], removed: ["\\Seen"], notApplied: [] });
      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("42", ["\\Flagged"], { uid: true });
      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("42", ["\\Seen"], { uid: true });
    });

    it("reports notApplied when the server silently drops a user keyword (Proton Bridge behaviour)", async () => {
      mockFetchOne
        .mockResolvedValueOnce({ uid: 5 }) // existence check
        .mockResolvedValueOnce({ uid: 5, flags: new Set(["\\Seen"]) }); // Proton dropped Important_Tag
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 5, ["Important_Tag"], []);

      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("5", ["Important_Tag"], { uid: true });
      expect(result).toEqual({ added: [], removed: [], notApplied: ["Important_Tag"] });
    });

    it("reports notApplied when a remove had no effect (flag still present)", async () => {
      mockFetchOne
        .mockResolvedValueOnce({ uid: 5 })
        .mockResolvedValueOnce({ uid: 5, flags: new Set(["StuckKeyword"]) });
      const service = new ImapService(baseConfig);
      const result = await service.updateFlags("INBOX", 5, [], ["StuckKeyword"]);

      expect(result).toEqual({ added: [], removed: [], notApplied: ["StuckKeyword"] });
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

  describe("findSentCopyMeta", () => {
    it("resolves \\Sent via specialUse and finds the sent copy on the first SEARCH attempt", async () => {
      mockList.mockResolvedValueOnce([{ path: "Sent", name: "Sent", specialUse: "\\Sent" }]);
      mockSearch.mockResolvedValueOnce([99]);
      mockFetchOne.mockResolvedValueOnce({
        uid: 99,
        headers: Buffer.from("Reply-To: me@pm.me\r\n"),
      });

      const service = new ImapService(baseConfig);
      const meta = await service.findSentCopyMeta("<sent@example.com>", { backoffMs: [0, 0, 0] });

      expect(meta.uid).toBe(99);
      expect(meta.folder).toBe("Sent");
      expect(meta.replyTo).toBe("me@pm.me");
    });

    it("retries SEARCH when the first attempt returns empty (Proton index lag)", async () => {
      mockList.mockResolvedValueOnce([{ path: "Sent", name: "Sent", specialUse: "\\Sent" }]);
      // First two attempts: empty (index lagging). Third: hit.
      mockSearch.mockResolvedValueOnce([]).mockResolvedValueOnce([]).mockResolvedValueOnce([101]);
      mockFetchOne.mockResolvedValueOnce({
        uid: 101,
        headers: Buffer.from("Reply-To: me@pm.me\r\n"),
      });

      const service = new ImapService(baseConfig);
      const meta = await service.findSentCopyMeta("<sent@example.com>", { backoffMs: [0, 0, 0] });

      expect(meta.uid).toBe(101);
      expect(mockSearch).toHaveBeenCalledTimes(3);
    });

    it("returns uid: undefined when retries exhaust without convergence", async () => {
      mockList.mockResolvedValueOnce([{ path: "Sent", name: "Sent", specialUse: "\\Sent" }]);
      mockSearch.mockResolvedValue([]); // never finds it

      const service = new ImapService(baseConfig);
      const meta = await service.findSentCopyMeta("<never@example.com>", { backoffMs: [0, 0, 0] });

      expect(meta.uid).toBeUndefined();
      expect(meta.folder).toBe("Sent");
    });

    it("falls back to the literal 'Sent' name when no \\Sent annotation is published", async () => {
      mockList.mockResolvedValueOnce([{ path: "INBOX", name: "Inbox", specialUse: "\\Inbox" }]);
      mockSearch.mockResolvedValueOnce([7]);
      mockFetchOne.mockResolvedValueOnce({ uid: 7, headers: Buffer.from("") });

      const service = new ImapService(baseConfig);
      const meta = await service.findSentCopyMeta("<x@y>", { backoffMs: [0, 0, 0] });

      expect(meta.folder).toBe("Sent");
      expect(meta.uid).toBe(7);
    });

    it("surfaces a rewritten Reply-To via the delivered header", async () => {
      mockList.mockResolvedValueOnce([{ path: "Sent", name: "Sent", specialUse: "\\Sent" }]);
      mockSearch.mockResolvedValueOnce([12]);
      // Requested replyTo was "someone-else@example.com"; Proton rewrote it.
      mockFetchOne.mockResolvedValueOnce({
        uid: 12,
        headers: Buffer.from("Reply-To: me@pm.me\r\n"),
      });

      const service = new ImapService(baseConfig);
      const meta = await service.findSentCopyMeta("<x@y>", { backoffMs: [0, 0, 0] });

      expect(meta.replyTo).toBe("me@pm.me");
    });

    it("returns replyTo: undefined when the delivered message has no Reply-To header (Proton's stripped-rewrite case)", async () => {
      // Proton normalizes Reply-To to match From by stripping the header
      // entirely. Caller-side rewrite-detection needs to treat that as a
      // signal — the absence of the header is itself the rewrite.
      mockList.mockResolvedValueOnce([{ path: "Sent", name: "Sent", specialUse: "\\Sent" }]);
      mockSearch.mockResolvedValueOnce([15]);
      mockFetchOne.mockResolvedValueOnce({
        uid: 15,
        headers: Buffer.from("Date: Mon, 25 May 2026 20:50:46 GMT\r\nSubject: hi\r\n"),
      });

      const service = new ImapService(baseConfig);
      const meta = await service.findSentCopyMeta("<x@y>", { backoffMs: [0, 0, 0] });

      expect(meta.uid).toBe(15);
      expect(meta.replyTo).toBeUndefined();
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

    it("routes its result through rerouteAllMailOrphans so All Mail-only members get a real folder + UID", async () => {
      // get_thread's default walk (INBOX + Sent + All Mail) surfaces a thread
      // stored in a user folder only via its All Mail virtual copy, pairing the
      // row with an All Mail UID that is invalid in any real folder. The read
      // path now reuses the same reroute the thread-mutation ops use. Here we
      // assert the delegation (and that its output is returned verbatim); the
      // reroute's orphan-rewrite behavior itself is covered by the previewThread
      // tests below.
      mockSearch.mockResolvedValueOnce([7]).mockResolvedValue([]); // seed in INBOX; walk refs empty
      mockFetchOne.mockResolvedValueOnce({
        uid: 7,
        envelope: {
          messageId: "<s@id>",
          inReplyTo: "",
          subject: "hi",
          from: [{ address: "a@b" }],
          to: [{ address: "c@d" }],
          date: new Date("2026-05-20T00:00:00Z"),
        },
        flags: new Set(),
        headers: Buffer.from(""),
      });
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 7,
            envelope: {
              messageId: "<s@id>",
              subject: "hi",
              from: [{ address: "a@b" }],
              to: [{ address: "c@d" }],
              date: new Date("2026-05-20T00:00:00Z"),
            },
            flags: new Set(),
          };
        })(),
      );

      const service = new ImapService(baseConfig);
      const rerouted = [
        {
          uid: 42,
          folder: "Folders/Development",
          messageId: "<s@id>",
          subject: "hi",
          from: "a@b",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
        },
      ];
      const spy = vi
        .spyOn(
          service as unknown as { rerouteAllMailOrphans: (m: unknown[]) => Promise<unknown[]> },
          "rerouteAllMailOrphans",
        )
        .mockResolvedValue(rerouted as never);

      const result = await service.getThreadByMessageId("<s@id>", 25);

      expect(spy).toHaveBeenCalledTimes(1);
      // Delegated the assembled thread rows…
      expect(spy).toHaveBeenCalledWith([expect.objectContaining({ folder: "INBOX", uid: 7, messageId: "<s@id>" })]);
      // …and returned the reroute output verbatim (real folder + UID, not All Mail).
      expect(result).toBe(rerouted);
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

    it("dedupes the same Message-ID across mailbox copies (INBOX + All Mail) and prefers the non-All-Mail row", async () => {
      // Proton's label model: one physical message can appear in INBOX *and* All Mail
      // with different UIDs but the same Message-ID. Headline count should be 1, not 2,
      // and the canonical row should be the INBOX copy with otherFolders=["All Mail"].
      const messageId = "<dup@example.com>";

      // Seed lookup: hit on INBOX (first folder in default walk).
      mockSearch
        .mockResolvedValueOnce([23]) // INBOX seed found
        // Reference walk: every header search returns its respective UID,
        // then empties for the rest. INBOX walks first (3 searches), then Sent (3, empty),
        // then All Mail (3 — first finds [33], rest empty).
        .mockResolvedValueOnce([23]) // INBOX Message-ID
        .mockResolvedValue([]); // remainder: empty

      // Seed envelope fetch (INBOX).
      mockFetchOne.mockResolvedValueOnce({
        uid: 23,
        envelope: {
          messageId,
          inReplyTo: "",
          subject: "shared",
          from: [{ address: "you@example.com" }],
          to: [{ address: "me@pm.me" }],
          date: new Date("2026-04-15T12:00:00Z"),
        },
        flags: new Set(),
        headers: Buffer.from(""),
      });

      // Override search for All Mail: when its 3 header searches run, the first returns [33].
      // The walk order is INBOX (3) → Sent (3) → All Mail (3). After INBOX's first hit we've used
      // 1 mockResolvedValueOnce above (seed) + 1 for INBOX Message-ID. Remaining searches default to [].
      // We need to re-arm so All Mail's first search returns [33].
      mockSearch
        .mockReset()
        .mockResolvedValueOnce([23]) // seed in INBOX
        .mockResolvedValueOnce([23]) // INBOX Message-ID search
        .mockResolvedValueOnce([]) // INBOX References
        .mockResolvedValueOnce([]) // INBOX In-Reply-To
        .mockResolvedValueOnce([]) // Sent Message-ID
        .mockResolvedValueOnce([]) // Sent References
        .mockResolvedValueOnce([]) // Sent In-Reply-To
        .mockResolvedValueOnce([33]) // All Mail Message-ID
        .mockResolvedValueOnce([]) // All Mail References
        .mockResolvedValueOnce([]); // All Mail In-Reply-To

      // Per-folder envelope fetch — both copies carry the same Message-ID.
      mockFetch
        .mockReturnValueOnce(
          (async function* () {
            yield {
              uid: 23,
              envelope: {
                messageId,
                subject: "shared",
                from: [{ address: "you@example.com" }],
                to: [{ address: "me@pm.me" }],
                date: new Date("2026-04-15T12:00:00Z"),
              },
              flags: new Set(),
            };
          })(),
        )
        .mockReturnValueOnce(
          (async function* () {
            yield {
              uid: 33,
              envelope: {
                messageId,
                subject: "shared",
                from: [{ address: "you@example.com" }],
                to: [{ address: "me@pm.me" }],
                date: new Date("2026-04-15T12:00:00Z"),
              },
              flags: new Set(),
            };
          })(),
        );

      const service = new ImapService(baseConfig);
      const thread = await service.getThreadByMessageId(messageId, 25);

      // One distinct message, presented via the INBOX copy.
      expect(thread.length).toBe(1);
      expect(thread[0].uid).toBe(23);
      const t = thread[0] as { folder?: string; mailboxCopies?: number; otherFolders?: string[] };
      expect(t.folder).toBe("INBOX");
      expect(t.mailboxCopies).toBe(2);
      expect(t.otherFolders).toEqual(["All Mail"]);
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
  // TRYCREATE signals into a uniform "Folder not found: <path>" error. An earlier text-matching
  // approach didn't actually catch the real imapflow errors; every call now routes
  // through a shared lockFolder() helper and this block pins that contract.
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

  describe("getSpecialFolders", () => {
    it("resolves special-use folder paths from the mailbox list", async () => {
      mockList.mockResolvedValue([
        { path: "INBOX" },
        { path: "Custom Trash Path", specialUse: "\\Trash" },
        { path: "Spam", specialUse: "\\Junk" },
      ]);

      const service = new ImapService(baseConfig);
      const result = await service.getSpecialFolders();

      expect(result.trash).toBe("Custom Trash Path");
      expect(result.junk).toBe("Spam");
    });

    it("memoizes across calls (single LIST per service instance)", async () => {
      mockList.mockResolvedValue([{ path: "Trash", specialUse: "\\Trash" }]);

      const service = new ImapService(baseConfig);
      await service.getSpecialFolders();
      await service.getSpecialFolders();
      await service.getSpecialFolders();

      expect(mockList).toHaveBeenCalledTimes(1);
    });
  });

  describe("moveMessage (explicit destination unchanged by resolver)", () => {
    it("explicit destination string is used as-is (resolver does not rewrite)", async () => {
      // Even if specialUse points elsewhere, calling moveMessage with the literal
      // "Trash" must go to "Trash". The resolver is only used by callers that
      // explicitly ask for the special-use mapping.
      mockList.mockResolvedValue([{ path: "Custom Trash", specialUse: "\\Trash" }]);
      mockFetchOne.mockResolvedValue({ uid: 7 });
      mockMessageMove.mockResolvedValue({ uidMap: new Map([[7, 42]]) });

      const service = new ImapService(baseConfig);
      await service.moveMessage("INBOX", 7, "Trash");

      expect(mockMessageMove).toHaveBeenCalledWith("7", "Trash", { uid: true });
    });
  });

  describe("resolveUidsFromCriteria", () => {
    it("returns UIDs matching the SEARCH criteria", async () => {
      mockSearch.mockResolvedValue([3, 5, 9]);

      const service = makeService();
      const uids = await service.resolveUidsFromCriteria("INBOX", { from: "alice@example.com" });

      expect(uids).toEqual([3, 5, 9]);
      expect(mockSearch).toHaveBeenCalledWith({ from: "alice@example.com" }, { uid: true });
    });

    it("returns [] when no matches", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      expect(await service.resolveUidsFromCriteria("INBOX", { from: "nobody" })).toEqual([]);
    });

    it("propagates List-Id and size filters into the SEARCH header object", async () => {
      mockSearch.mockResolvedValue([1]);
      const service = makeService();
      await service.resolveUidsFromCriteria("INBOX", { listId: "substack.com", larger: 5000 });
      expect(mockSearch).toHaveBeenCalledWith(
        expect.objectContaining({ larger: 5000, header: { "List-Id": "substack.com" } }),
        { uid: true },
      );
    });

    it("throws if hasAttachment is in criteria", async () => {
      const service = makeService();
      await expect(service.resolveUidsFromCriteria("INBOX", { hasAttachment: true })).rejects.toThrow(/hasAttachment/);
    });
  });

  describe("bulkMove", () => {
    it("moves the given UIDs, reports newUids from UIDPLUS map", async () => {
      mockFetch.mockImplementation(() => uidIter([1, 2, 3])); // pre-check FETCH: all exist
      mockMessageMove.mockResolvedValue({
        uidMap: new Map([
          [1, 101],
          [2, 102],
          [3, 103],
        ]),
      });

      const service = makeService();
      const result = await service.bulkMove("INBOX", [1, 2, 3], "Archive");

      expect(result.moved).toBe(3);
      expect(result.notFound).toEqual([]);
      expect(result.newUids).toEqual([101, 102, 103]);
      expect(result.destination).toBe("Archive");
      expect(mockMessageMove).toHaveBeenCalledWith("1,2,3", "Archive", { uid: true });
    });

    it("reports notFound for UIDs the pre-check FETCH did not yield", async () => {
      mockFetch.mockImplementation(() => uidIter([1, 3])); // 2 is phantom
      mockMessageMove.mockResolvedValue({
        uidMap: new Map([
          [1, 101],
          [3, 103],
        ]),
      });

      const service = makeService();
      const result = await service.bulkMove("INBOX", [1, 2, 3], "Archive");

      expect(result.moved).toBe(2);
      expect(result.notFound).toEqual([2]);
      expect(mockMessageMove).toHaveBeenCalledWith("1,3", "Archive", { uid: true });
    });

    it("when all UIDs are phantom, does not call messageMove and returns moved=0", async () => {
      mockFetch.mockImplementation(() => uidIter([]));
      const service = makeService();
      const result = await service.bulkMove("INBOX", [1, 2], "Archive");

      expect(result.moved).toBe(0);
      expect(result.notFound).toEqual([1, 2]);
      expect(mockMessageMove).not.toHaveBeenCalled();
    });

    it("translates missing destination into 'Destination folder not found' error", async () => {
      mockFetch.mockImplementation(() => uidIter([1]));
      mockMessageMove.mockResolvedValue(false);
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));

      const service = makeService();
      await expect(service.bulkMove("INBOX", [1], "Bogus")).rejects.toThrow(/Destination folder not found: Bogus/);
    });
  });

  describe("bulkDelete", () => {
    it("permanent=true expunges via messageDelete", async () => {
      mockFetch.mockImplementation(() => uidIter([1, 2]));
      mockMessageDelete.mockResolvedValue(true);

      const service = makeService();
      const result = await service.bulkDelete("INBOX", [1, 2], true);

      expect(result.deleted).toBe(2);
      expect(result.notFound).toEqual([]);
      expect(result.newUids).toBeUndefined();
      expect(mockMessageDelete).toHaveBeenCalledWith("1,2", { uid: true });
    });

    it("permanent=false moves to resolved Trash path and reports newUids", async () => {
      mockList.mockResolvedValue([{ path: "Trash Bin", specialUse: "\\Trash" }]);
      mockFetch.mockImplementation(() => uidIter([1, 2]));
      mockMessageMove.mockResolvedValue({
        uidMap: new Map([
          [1, 201],
          [2, 202],
        ]),
      });

      const service = makeService();
      const result = await service.bulkDelete("INBOX", [1, 2], false);

      expect(result.deleted).toBe(2);
      expect(result.destination).toBe("Trash Bin");
      expect(result.newUids).toEqual([201, 202]);
      expect(mockMessageMove).toHaveBeenCalledWith("1,2", "Trash Bin", { uid: true });
    });

    it("falls back to literal 'Trash' when special-use is not annotated", async () => {
      mockList.mockResolvedValue([{ path: "INBOX" }]); // no \Trash mailbox
      mockFetch.mockImplementation(() => uidIter([1]));
      mockMessageMove.mockResolvedValue({ uidMap: new Map() });

      const service = makeService();
      const result = await service.bulkDelete("INBOX", [1], false);

      expect(result.destination).toBe("Trash");
    });

    it("reports notFound UIDs from the pre-check", async () => {
      mockFetch.mockImplementation(() => uidIter([1])); // 99 is phantom
      mockMessageDelete.mockResolvedValue(true);

      const service = makeService();
      const result = await service.bulkDelete("INBOX", [1, 99], true);

      expect(result.notFound).toEqual([99]);
      expect(mockMessageDelete).toHaveBeenCalledWith("1", { uid: true });
    });

    it("trusts FETCH over SEARCH for the existence pre-check (regression: Proton SEARCH lag)", async () => {
      // Reproduces a Proton Mail Bridge quirk: IMAP SEARCH can
      // lag the live mailbox state. If pre-check trusted SEARCH, two UIDs would be
      // (incorrectly) reported as notFound even though FETCH (= ground truth)
      // confirms all three exist and would be deleted successfully. This locks in
      // the contract that the pre-check uses FETCH, not SEARCH.
      mockSearch.mockResolvedValue([7]); // stale index — disagrees with reality
      mockFetch.mockImplementation(() => uidIter([5, 6, 7])); // ground truth
      mockMessageDelete.mockResolvedValue(true);

      const service = makeService();
      const result = await service.bulkDelete("Sent", [5, 6, 7], true);

      expect(result.deleted).toBe(3);
      expect(result.notFound).toEqual([]);
      expect(mockMessageDelete).toHaveBeenCalledWith("5,6,7", { uid: true });
    });
  });

  describe("bulkUpdateFlags", () => {
    it("adds and removes flags in one connection", async () => {
      // First FETCH: existence pre-check. Second FETCH: post-STORE verify.
      mockFetch
        .mockImplementationOnce(() => uidIter([1, 2, 3]))
        .mockImplementationOnce(() =>
          uidIterWithFlags([
            [1, ["\\Seen"]],
            [2, ["\\Seen"]],
            [3, ["\\Seen"]],
          ]),
        );
      mockMessageFlagsAdd.mockResolvedValue(true);
      mockMessageFlagsRemove.mockResolvedValue(true);

      const service = makeService();
      const result = await service.bulkUpdateFlags("INBOX", [1, 2, 3], ["\\Seen"], ["\\Flagged"]);

      expect(result.affected).toBe(3);
      expect(result.notApplied).toEqual([]);
      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("1,2,3", ["\\Seen"], { uid: true });
      expect(mockMessageFlagsRemove).toHaveBeenCalledWith("1,2,3", ["\\Flagged"], { uid: true });
    });

    it("skips messageFlagsAdd when add list is empty", async () => {
      mockFetch.mockImplementationOnce(() => uidIter([1])).mockImplementationOnce(() => uidIterWithFlags([[1, []]])); // \Seen removed
      const service = makeService();
      const result = await service.bulkUpdateFlags("INBOX", [1], [], ["\\Seen"]);
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
      expect(mockMessageFlagsRemove).toHaveBeenCalled();
      expect(result.notApplied).toEqual([]);
    });

    it("reports notFound from pre-check", async () => {
      mockFetch
        .mockImplementationOnce(() => uidIter([1]))
        .mockImplementationOnce(() => uidIterWithFlags([[1, ["\\Seen"]]]));
      const service = makeService();
      const result = await service.bulkUpdateFlags("INBOX", [1, 99], ["\\Seen"], []);
      expect(result.notFound).toEqual([99]);
      expect(mockMessageFlagsAdd).toHaveBeenCalledWith("1", ["\\Seen"], { uid: true });
    });

    it("reports notApplied when the server silently drops a user keyword on every UID", async () => {
      mockFetch
        .mockImplementationOnce(() => uidIter([1, 2]))
        .mockImplementationOnce(() =>
          uidIterWithFlags([
            [1, ["\\Seen"]], // Important_Tag absent — dropped
            [2, ["\\Seen"]],
          ]),
        );
      const service = makeService();
      const result = await service.bulkUpdateFlags("INBOX", [1, 2], ["Important_Tag"], []);
      expect(result.affected).toBe(2);
      expect(result.notApplied).toEqual(["Important_Tag"]);
    });
  });

  describe("bulkUpdateLabels", () => {
    it("adds a label to many UIDs via one batched COPY per label after status() pre-check", async () => {
      mockFetch
        .mockImplementationOnce(() => uidIter([1, 2, 3])) // existence check
        .mockImplementationOnce(() =>
          (async function* () {
            yield { uid: 1, envelope: { messageId: "<m1@x>" } };
            yield { uid: 2, envelope: { messageId: "<m2@x>" } };
            yield { uid: 3, envelope: { messageId: "<m3@x>" } };
          })(),
        ); // Message-ID gather
      mockStatus.mockResolvedValue({ messages: 0 }); // label exists
      mockMessageCopy.mockResolvedValue({ destination: "Labels/Important" });

      const service = makeService();
      const result = await service.bulkUpdateLabels("INBOX", [1, 2, 3], ["Labels/Important"], []);

      expect(result.affected).toBe(3);
      expect(result.notFound).toEqual([]);
      expect(result.notApplied).toEqual([]);
      expect(mockMessageCopy).toHaveBeenCalledWith("1,2,3", "Labels/Important", { uid: true });
    });

    it("throws 'Label not found' when status() pre-check fails for an add target", async () => {
      mockFetch
        .mockImplementationOnce(() => uidIter([1]))
        .mockImplementationOnce(() =>
          (async function* () {
            yield { uid: 1, envelope: { messageId: "<m1@x>" } };
          })(),
        );
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));

      const service = makeService();
      await expect(service.bulkUpdateLabels("INBOX", [1], ["Labels/Missing"], [])).rejects.toThrow(
        /Label not found: Labels\/Missing/,
      );
      expect(mockMessageCopy).not.toHaveBeenCalled();
    });

    it("rejects label paths that do not start with 'Labels/'", async () => {
      const service = makeService();
      await expect(service.bulkUpdateLabels("INBOX", [1], ["NotALabel"], [])).rejects.toThrow(
        /must start with "Labels\/"/,
      );
    });

    it("reports notApplied when a remove targets a label that doesn't exist as a mailbox", async () => {
      mockFetch
        .mockImplementationOnce(() => uidIter([1]))
        .mockImplementationOnce(() =>
          (async function* () {
            yield { uid: 1, envelope: { messageId: "<m1@x>" } };
          })(),
        );
      const originalLock = mockGetMailboxLock.getMockImplementation();
      mockGetMailboxLock.mockImplementation((path: string) => {
        if (path === "Labels/Stale") {
          return Promise.reject(Object.assign(new Error("nope"), { mailboxMissing: true }));
        }
        return Promise.resolve({ release: vi.fn() });
      });

      const service = makeService();
      const result = await service.bulkUpdateLabels("INBOX", [1], [], ["Labels/Stale"]);

      expect(result.notApplied).toEqual(["Labels/Stale"]);
      mockGetMailboxLock.mockImplementation(originalLock ?? (() => Promise.resolve({ release: vi.fn() })));
    });

    it("returns affected=0 + populated notFound when every UID is a phantom", async () => {
      mockFetch.mockImplementationOnce(() => uidIter([]));
      const service = makeService();
      const result = await service.bulkUpdateLabels("INBOX", [1, 2], ["Labels/X"], []);
      expect(result.affected).toBe(0);
      expect(result.notFound).toEqual([1, 2]);
      expect(mockMessageCopy).not.toHaveBeenCalled();
    });

    it("locks the source folder BEFORE the existence pre-check (regression: existingUidSet ran without a selected mailbox)", async () => {
      // Track the lock state when each operation runs. The bug was that
      // existingUidSet (= a FETCH) executed outside the source-folder lock,
      // which on a fresh connection means "no mailbox selected" — FETCH then
      // returns nothing and every UID is reported as notFound while the
      // subsequent COPY still succeeds.
      let lockedFolder: string | null = null;
      const release = vi.fn(() => {
        lockedFolder = null;
      });
      const originalLock = mockGetMailboxLock.getMockImplementation();
      mockGetMailboxLock.mockImplementation((path: string) => {
        lockedFolder = path;
        return Promise.resolve({ release });
      });

      const fetchCallsLocked: (string | null)[] = [];
      mockFetch.mockImplementation(() => {
        fetchCallsLocked.push(lockedFolder);
        return uidIter([6]);
      });
      mockStatus.mockResolvedValue({ messages: 0 });
      mockMessageCopy.mockResolvedValue({ destination: "Labels/mcp-probe" });

      const service = makeService();
      const result = await service.bulkUpdateLabels("INBOX", [6], ["Labels/mcp-probe"], []);

      // Every FETCH during the operation must have happened with INBOX locked.
      expect(fetchCallsLocked.length).toBeGreaterThan(0);
      for (const lf of fetchCallsLocked) {
        expect(lf).toBe("INBOX");
      }
      // And the UID is correctly identified as present, not as notFound.
      expect(result.affected).toBe(1);
      expect(result.notFound).toEqual([]);

      mockGetMailboxLock.mockImplementation(originalLock ?? (() => Promise.resolve({ release: vi.fn() })));
    });
  });

  describe("filterExistingUids (dry-run accounting parity)", () => {
    it("partitions a UID list into existing vs notFound via the same FETCH used by live ops", async () => {
      mockFetch.mockImplementation(() => uidIter([5])); // 999999 is phantom — only 5 exists
      const service = makeService();
      const result = await service.filterExistingUids("INBOX", [5, 999999]);
      // Order preserved from the input — the agent-facing UI iterates the input
      // list, so callers can trust the position of each UID.
      expect(result).toEqual({ existing: [5], notFound: [999999] });
    });

    it("returns empty partitions on an empty input without contacting the server", async () => {
      const service = makeService();
      const result = await service.filterExistingUids("INBOX", []);
      expect(result).toEqual({ existing: [], notFound: [] });
      expect(mockFetch).not.toHaveBeenCalled();
      expect(mockGetMailboxLock).not.toHaveBeenCalled();
    });

    it("returns all uids as notFound when none exist", async () => {
      mockFetch.mockImplementation(() => uidIter([]));
      const service = makeService();
      const result = await service.filterExistingUids("INBOX", [10, 20, 30]);
      expect(result.existing).toEqual([]);
      expect(result.notFound).toEqual([10, 20, 30]);
    });
  });

  describe("resolveAndFilterUidsFromCriteria (B-04 phantom-UID guard)", () => {
    it("filters out phantom UIDs that SEARCH still surfaces after a move/delete", async () => {
      // Simulate Proton Bridge SEARCH lag: SEARCH reports [1, 2, 3] but FETCH
      // (used by live ops) only confirms [2, 3] — UID 1 has been moved away
      // and the SEARCH index hasn't caught up. The dry-run preview must
      // match the FETCH-authoritative view of live bulk ops.
      mockSearch.mockResolvedValue([1, 2, 3]);
      mockFetch.mockImplementation(() => uidIter([2, 3]));
      const service = makeService();
      const result = await service.resolveAndFilterUidsFromCriteria("INBOX", { from: "alice@example.com" });
      expect(result).toEqual([2, 3]);
    });

    it("short-circuits without a FETCH when SEARCH returns empty", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      const result = await service.resolveAndFilterUidsFromCriteria("INBOX", { from: "nobody" });
      expect(result).toEqual([]);
      expect(mockFetch).not.toHaveBeenCalled();
    });

    it("returns the full SEARCH set when every UID survives the FETCH pass", async () => {
      mockSearch.mockResolvedValue([10, 20]);
      mockFetch.mockImplementation(() => uidIter([10, 20]));
      const service = makeService();
      const result = await service.resolveAndFilterUidsFromCriteria("INBOX", { from: "alice@example.com" });
      expect(result).toEqual([10, 20]);
    });
  });

  describe("createFolder", () => {
    it("creates a new folder", async () => {
      mockMailboxCreate.mockResolvedValue({ path: "Receipts", created: true });
      const service = makeService();
      const result = await service.createFolder("Receipts");
      expect(result).toEqual({ created: true });
      expect(mockMailboxCreate).toHaveBeenCalledWith("Receipts");
    });

    it("returns alreadyExists:true when imapflow reports the folder exists", async () => {
      mockMailboxCreate.mockRejectedValue(
        Object.assign(new Error("Mailbox already exists"), { serverResponseCode: "ALREADYEXISTS" }),
      );
      const service = makeService();
      const result = await service.createFolder("Existing");
      expect(result).toEqual({ created: false, alreadyExists: true });
    });

    it("translates root-namespace failures into a Proton-specific hint", async () => {
      mockMailboxCreate.mockRejectedValue(new Error("Command failed"));
      // status() probe must report 'mailbox missing' so the namespace hint path is hit.
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      const service = makeService();
      await expect(service.createFolder("Receipts")).rejects.toThrow(
        /folders must be created under the "Folders\/" namespace.*Folders\/Receipts/,
      );
    });

    it("wraps an opaque failure (no collision in list) into an actionable message that still surfaces the cause", async () => {
      mockMailboxCreate.mockRejectedValue(new Error("Some other failure"));
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      mockList.mockResolvedValue([{ path: "INBOX" }, { path: "Folders/Other" }]); // no leaf collision
      const service = makeService();
      await expect(service.createFolder("Folders/Receipts")).rejects.toThrow(
        /Proton rejected the request.*Some other failure/s,
      );
    });

    it("treats bare 'Command failed' as idempotent when the folder actually exists (Proton case)", async () => {
      // Proton Mail Bridge returns a bare 'Command failed' with no ALREADYEXISTS code
      // when re-creating an existing folder. The status() probe disambiguates.
      mockMailboxCreate.mockRejectedValue(new Error("Command failed"));
      mockStatus.mockResolvedValue({ messages: 0 });
      const service = makeService();
      const result = await service.createFolder("Folders/MCP_Test");
      expect(result).toEqual({ created: false, alreadyExists: true });
    });

    it("rejects path-traversal segments (symmetry with delete_folder)", async () => {
      const service = makeService();
      await expect(service.createFolder("Folders/../Escape")).rejects.toThrow(/invalid segments/);
      await expect(service.createFolder("Folders/./Sub")).rejects.toThrow(/invalid segments/);
      expect(mockMailboxCreate).not.toHaveBeenCalled();
    });

    it("redirects Labels/ paths to create_label", async () => {
      const service = makeService();
      await expect(service.createFolder("Labels/SneakyAsLabel")).rejects.toThrow(/Use the create_label tool/);
      expect(mockMailboxCreate).not.toHaveBeenCalled();
    });
  });

  describe("createLabel", () => {
    it("creates a label via the dedicated path (regression: create_label was broken — Labels/ guard recursed)", async () => {
      mockMailboxCreate.mockResolvedValue({ path: "Labels/Receipts", created: true });
      const service = makeService();
      const result = await service.createLabel("Receipts");
      expect(result).toEqual({ created: true, path: "Labels/Receipts" });
      expect(mockMailboxCreate).toHaveBeenCalledWith("Labels/Receipts");
    });

    it("is idempotent: returns alreadyExists:true on a second create", async () => {
      mockMailboxCreate.mockRejectedValue(
        Object.assign(new Error("Mailbox already exists"), { serverResponseCode: "ALREADYEXISTS" }),
      );
      const service = makeService();
      const result = await service.createLabel("Important");
      expect(result).toEqual({ created: false, alreadyExists: true, path: "Labels/Important" });
    });

    it("rejects names containing '/'", async () => {
      const service = makeService();
      await expect(service.createLabel("Bad/Name")).rejects.toThrow(/bare name without "\/"/);
      expect(mockMailboxCreate).not.toHaveBeenCalled();
    });

    it("explains a name collision when Proton refuses with a bare 'Command failed' (the Labels/Archive case)", async () => {
      // create_label("Archive") fails with an opaque "Command failed"
      // while Important/Test succeed. Proton enforces unique names across
      // folders and labels, so a label can't reuse the system "Archive" folder
      // name. The status() probe finds nothing (the label wasn't created), and
      // the live list() shows the colliding mailbox — so we explain it instead
      // of leaking "Command failed".
      mockMailboxCreate.mockRejectedValue(new Error("Command failed"));
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      mockList.mockResolvedValue([
        { path: "INBOX" },
        { path: "Archive", specialUse: "\\Archive" },
        { path: "Labels/Important" },
      ]);
      const service = makeService();
      await expect(service.createLabel("Archive")).rejects.toThrow(
        /a mailbox named "Archive" already exists \("Archive"\).*reuse a name held by another mailbox/s,
      );
    });

    it("treats a bare 'Command failed' as idempotent when list() shows the label already exists (probe-miss safety net)", async () => {
      // If status() is unreliable on a Proton label-backed mailbox, the probe
      // can wrongly report "missing" for a label that already exists. The
      // list() check catches the exact-path match and returns idempotent
      // success rather than erroring.
      mockMailboxCreate.mockRejectedValue(new Error("Command failed"));
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      mockList.mockResolvedValue([{ path: "INBOX" }, { path: "Labels/Important" }]);
      const service = makeService();
      const result = await service.createLabel("Important");
      expect(result).toEqual({ created: false, alreadyExists: true, path: "Labels/Important" });
    });

    it("rejects empty names", async () => {
      const service = makeService();
      await expect(service.createLabel("")).rejects.toThrow(/non-empty bare name/);
    });
  });

  describe("renameFolder", () => {
    it("renames a folder under Folders/", async () => {
      mockMailboxRename.mockResolvedValue({ newPath: "Folders/Renamed" });
      const service = makeService();
      await service.renameFolder("Folders/Old", "Folders/Renamed");
      expect(mockMailboxRename).toHaveBeenCalledWith("Folders/Old", "Folders/Renamed");
    });

    it("renames a label container under Labels/", async () => {
      mockMailboxRename.mockResolvedValue({ newPath: "Labels/New" });
      const service = makeService();
      await service.renameFolder("Labels/Old", "Labels/New");
      expect(mockMailboxRename).toHaveBeenCalledWith("Labels/Old", "Labels/New");
    });

    it("rejects renaming a system mailbox (INBOX)", async () => {
      const service = makeService();
      await expect(service.renameFolder("INBOX", "Folders/Hacked")).rejects.toThrow(
        /restricted to the "Folders\/" and "Labels\/"/,
      );
      await expect(service.renameFolder("Sent", "Folders/X")).rejects.toThrow(
        /restricted to the "Folders\/" and "Labels\/"/,
      );
      await expect(service.renameFolder("Trash", "Folders/X")).rejects.toThrow(
        /restricted to the "Folders\/" and "Labels\/"/,
      );
      expect(mockMailboxRename).not.toHaveBeenCalled();
    });

    it("rejects renaming TO a system-mailbox path (no namespace prefix)", async () => {
      const service = makeService();
      await expect(service.renameFolder("Folders/Real", "INBOX")).rejects.toThrow(
        /restricted to the "Folders\/" and "Labels\/"/,
      );
      expect(mockMailboxRename).not.toHaveBeenCalled();
    });

    it("rejects path-traversal segments on either end", async () => {
      const service = makeService();
      await expect(service.renameFolder("Folders/../INBOX", "Folders/X")).rejects.toThrow(/invalid segments/);
      await expect(service.renameFolder("Folders/A", "Folders/../INBOX")).rejects.toThrow(/invalid segments/);
      expect(mockMailboxRename).not.toHaveBeenCalled();
    });

    it("translates missing source into 'Folder not found' error", async () => {
      mockMailboxRename.mockRejectedValue(Object.assign(new Error("Mailbox doesn't exist"), { mailboxMissing: true }));
      const service = makeService();
      await expect(service.renameFolder("Folders/Missing", "Folders/X")).rejects.toThrow(
        /Folder not found: Folders\/Missing/,
      );
    });

    it("falls back to status() probe when Proton returns bare 'Command failed' (no mailboxMissing signal)", async () => {
      // Proton Mail Bridge returns 'Command failed' without annotating the error.
      // The post-failure probe re-runs as status(from) and surfaces mailboxMissing.
      mockMailboxRename.mockRejectedValue(new Error("Command failed"));
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      const service = makeService();
      await expect(service.renameFolder("Folders/DoesNotExist", "Folders/Nope")).rejects.toThrow(
        /Folder not found: Folders\/DoesNotExist/,
      );
    });

    it("rethrows the original error when the probe shows the source exists", async () => {
      // Genuine other failure (e.g. server transient): probe succeeds so we don't
      // pretend it was a missing-folder error.
      mockMailboxRename.mockRejectedValue(new Error("server hiccup"));
      mockStatus.mockResolvedValue({ messages: 0 });
      const service = makeService();
      await expect(service.renameFolder("Folders/Real", "Folders/X")).rejects.toThrow(/server hiccup/);
    });
  });

  describe("deleteFolder", () => {
    it("deletes a folder under Folders/", async () => {
      mockMailboxDelete.mockResolvedValue(undefined);
      const service = makeService();
      await service.deleteFolder("Folders/Old");
      expect(mockMailboxDelete).toHaveBeenCalledWith("Folders/Old");
    });

    it("deletes a label container under Labels/", async () => {
      mockMailboxDelete.mockResolvedValue(undefined);
      const service = makeService();
      await service.deleteFolder("Labels/Archived");
      expect(mockMailboxDelete).toHaveBeenCalledWith("Labels/Archived");
    });

    it("rejects paths outside Folders/ and Labels/ namespaces", async () => {
      const service = makeService();
      await expect(service.deleteFolder("INBOX")).rejects.toThrow(/restricted to the "Folders\/" and "Labels\/"/);
      await expect(service.deleteFolder("Trash")).rejects.toThrow(/restricted to the "Folders\/" and "Labels\/"/);
      expect(mockMailboxDelete).not.toHaveBeenCalled();
    });

    it("translates annotated missing-folder errors into 'Folder not found'", async () => {
      mockMailboxDelete.mockRejectedValue(Object.assign(new Error("Mailbox doesn't exist"), { mailboxMissing: true }));
      const service = makeService();
      await expect(service.deleteFolder("Folders/Missing")).rejects.toThrow(/Folder not found: Folders\/Missing/);
    });

    it("probes with status() when Proton returns bare 'Command failed' (no annotation)", async () => {
      mockMailboxDelete.mockRejectedValue(new Error("Command failed"));
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      const service = makeService();
      await expect(service.deleteFolder("Folders/Ghost")).rejects.toThrow(/Folder not found: Folders\/Ghost/);
    });

    it("DOES accept `.`/`..` segments under Folders/Labels — needed for cleanup of adversarial paths", async () => {
      // The segment guard on delete was dropped in v1.0.0 follow-up: IMAP treats
      // paths as opaque literal names (no parent-dir semantics), so the guard
      // gave no security but blocked legitimate cleanup of paths created by
      // other IMAP clients (a real agent-report mailbox state). Same guard
      // remains on create/rename, which is where the value lives.
      mockMailboxDelete.mockResolvedValue(undefined);
      const service = makeService();
      await service.deleteFolder("Folders/../Escape");
      expect(mockMailboxDelete).toHaveBeenCalledWith("Folders/../Escape");
    });

    it("reports the relocated message count for a non-empty folder (Proton moves contents to All Mail)", async () => {
      mockList.mockResolvedValueOnce([]); // no nested children
      mockStatus.mockResolvedValueOnce({ messages: 7 }); // 7 messages about to relocate
      mockMailboxDelete.mockResolvedValueOnce(undefined);
      const service = makeService();
      const result = await service.deleteFolder("Folders/Archive");
      expect(result).toEqual({ children: [], messageCount: 7, isLabel: false });
    });

    it("flags a label deletion (isLabel) so the response can say 'tag removed', not 'relocated'", async () => {
      mockList.mockResolvedValueOnce([]);
      mockStatus.mockResolvedValueOnce({ messages: 3 });
      mockMailboxDelete.mockResolvedValueOnce(undefined);
      const service = makeService();
      const result = await service.deleteFolder("Labels/Receipts");
      expect(result).toEqual({ children: [], messageCount: 3, isLabel: true });
    });
  });

  describe("updateLabels", () => {
    it("adds labels via messageCopy after a status() pre-check confirms each label exists", async () => {
      mockFetchOne.mockResolvedValue({ uid: 7, envelope: { messageId: "<m@id>" } });
      mockStatus.mockResolvedValue({ messages: 0 }); // pre-check: both labels exist
      mockMessageCopy.mockResolvedValue({ destination: "Labels/X" });

      const service = makeService();
      const result = await service.updateLabels("INBOX", 7, ["Labels/Important", "Labels/Work"], []);

      expect(result).toEqual({ added: ["Labels/Important", "Labels/Work"], removed: [], notApplied: [] });
      expect(mockStatus).toHaveBeenNthCalledWith(1, "Labels/Important", { messages: true });
      expect(mockStatus).toHaveBeenNthCalledWith(2, "Labels/Work", { messages: true });
      expect(mockMessageCopy).toHaveBeenNthCalledWith(1, "7", "Labels/Important", { uid: true });
      expect(mockMessageCopy).toHaveBeenNthCalledWith(2, "7", "Labels/Work", { uid: true });
      expect(mockMessageDelete).not.toHaveBeenCalled();
    });

    it("throws 'Label not found' when the pre-check status() rejects with mailboxMissing", async () => {
      mockFetchOne.mockResolvedValue({ uid: 7, envelope: { messageId: "<m@id>" } });
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      const service = makeService();
      await expect(service.updateLabels("INBOX", 7, ["Labels/DoesNotExist"], [])).rejects.toThrow(
        /Label not found: Labels\/DoesNotExist/,
      );
      // messageCopy must NOT be reached when the pre-check fails — that was the v0.6.0 bug.
      expect(mockMessageCopy).not.toHaveBeenCalled();
    });

    it("throws 'Label not found' even when messageCopy would return falsy silently (Proton case)", async () => {
      // Reproduces the BUG 1 conditions exactly: Proton Bridge returns no
      // mailboxMissing signal AND imapflow's messageCopy resolves false rather
      // than throwing. The pre-check is what makes this case loud.
      mockFetchOne.mockResolvedValue({ uid: 7, envelope: { messageId: "<m@id>" } });
      mockStatus.mockRejectedValue(Object.assign(new Error("nope"), { mailboxMissing: true }));
      mockMessageCopy.mockResolvedValue(false); // would silently succeed without the pre-check
      const service = makeService();
      await expect(service.updateLabels("INBOX", 7, ["Labels/DoesNotExist"], [])).rejects.toThrow(
        /Label not found: Labels\/DoesNotExist/,
      );
    });

    it("removes labels by Message-ID + messageDelete in the label mailbox", async () => {
      mockFetchOne.mockResolvedValue({ uid: 7, envelope: { messageId: "<m@id>" } });
      mockSearch.mockResolvedValue([42]); // UID in the label mailbox
      mockMessageDelete.mockResolvedValue(true);

      const service = makeService();
      const result = await service.updateLabels("INBOX", 7, [], ["Labels/Old"]);

      expect(result).toEqual({ added: [], removed: ["Labels/Old"], notApplied: [] });
      expect(mockSearch).toHaveBeenCalledWith({ header: { "Message-ID": "<m@id>" } }, { uid: true });
      expect(mockMessageDelete).toHaveBeenCalledWith("42", { uid: true });
    });

    it("reports labels the message doesn't carry under notApplied (no-op, not 'removed')", async () => {
      mockFetchOne.mockResolvedValue({ uid: 7, envelope: { messageId: "<m@id>" } });
      mockSearch.mockResolvedValue([]); // Message-ID not in this label mailbox

      const service = makeService();
      const result = await service.updateLabels("INBOX", 7, [], ["Labels/NotApplied"]);

      expect(result).toEqual({ added: [], removed: [], notApplied: ["Labels/NotApplied"] });
      expect(mockMessageDelete).not.toHaveBeenCalled();
    });

    it("reports missing label mailboxes under notApplied (idempotency)", async () => {
      mockFetchOne.mockResolvedValue({ uid: 7, envelope: { messageId: "<m@id>" } });
      mockGetMailboxLock
        .mockResolvedValueOnce({ release: vi.fn() }) // source folder INBOX
        .mockRejectedValueOnce(Object.assign(new Error("nope"), { mailboxMissing: true })); // label

      const service = makeService();
      const result = await service.updateLabels("INBOX", 7, [], ["Labels/NeverExisted"]);

      expect(result).toEqual({ added: [], removed: [], notApplied: ["Labels/NeverExisted"] });
      expect(mockMessageDelete).not.toHaveBeenCalled();
    });

    it("rejects label paths that don't start with Labels/", async () => {
      const service = makeService();
      await expect(service.updateLabels("INBOX", 7, ["Important"], [])).rejects.toThrow(
        /Label paths must start with "Labels\/"/,
      );
      await expect(service.updateLabels("INBOX", 7, [], ["Folders/Receipts"])).rejects.toThrow(
        /Label paths must start with "Labels\/"/,
      );
    });

    it("throws when the source UID is not found", async () => {
      mockFetchOne.mockResolvedValue(null);
      const service = makeService();
      await expect(service.updateLabels("INBOX", 999, ["Labels/Important"], [])).rejects.toThrow(
        /Message UID 999 not found in INBOX/,
      );
    });
  });

  describe("emptyFolder", () => {
    it("deletes all messages atomically via messageDelete; returns expunged count", async () => {
      mockSearch.mockResolvedValue([10, 11, 12]);
      mockMessageDelete.mockResolvedValue(true);

      const service = makeService();
      const result = await service.emptyFolder("Trash");

      expect(result.expunged).toBe(3);
      expect(mockMessageDelete).toHaveBeenCalledWith("10,11,12", { uid: true });
      // Old code path used messageFlagsAdd + a non-existent expunge() — neither should fire now.
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
    });

    it("returns {expunged: 0} for empty folder without calling messageDelete", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      const result = await service.emptyFolder("Trash");
      expect(result.expunged).toBe(0);
      expect(mockMessageDelete).not.toHaveBeenCalled();
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
    });

    it("dryRun reports the count that would be deleted WITHOUT calling messageDelete", async () => {
      mockSearch.mockResolvedValue([10, 11, 12]);
      const service = makeService();
      const result = await service.emptyFolder("Trash", { dryRun: true });
      expect(result.expunged).toBe(3);
      expect(mockMessageDelete).not.toHaveBeenCalled();
      expect(mockMessageFlagsAdd).not.toHaveBeenCalled();
    });

    it("dryRun on an empty folder returns 0 and deletes nothing", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      const result = await service.emptyFolder("Trash", { dryRun: true });
      expect(result.expunged).toBe(0);
      expect(mockMessageDelete).not.toHaveBeenCalled();
    });
  });

  describe("searchMessages — new filters", () => {
    it("passes larger and smaller through to client.search", async () => {
      mockSearch.mockResolvedValue([1]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield { uid: 1, envelope: {}, flags: [] };
        })(),
      );

      const service = makeService();
      await service.searchMessages("INBOX", { larger: 5000, smaller: 1_000_000 }, 10);
      expect(mockSearch).toHaveBeenCalledWith(expect.objectContaining({ larger: 5000, smaller: 1_000_000 }), {
        uid: true,
      });
    });

    it("passes listId as a header search", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      await service.searchMessages("INBOX", { listId: "politics.substack.com" }, 10);
      expect(mockSearch).toHaveBeenCalledWith(
        expect.objectContaining({ header: { "List-Id": "politics.substack.com" } }),
        { uid: true },
      );
    });

    it("hasAttachment=true post-filters via bodyStructure", async () => {
      mockSearch.mockResolvedValue([1, 2]);
      // Message 1 has an attachment, message 2 does not.
      mockFetch.mockReturnValue(
        (async function* () {
          yield { uid: 1, envelope: {}, flags: [], bodyStructure: { disposition: "attachment", part: "2" } };
          yield { uid: 2, envelope: {}, flags: [], bodyStructure: { type: "text/plain", part: "1" } };
        })(),
      );

      const service = makeService();
      const result = await service.searchMessages("INBOX", { hasAttachment: true }, 10);
      expect(result.map((m) => m.uid)).toEqual([1]);
    });

    it("hasAttachment=true excludes inline images from the match set", async () => {
      // Newsletter with an inline banner image and a non-attachment text part.
      // The old behavior would (mistakenly) report this as hasAttachment=true
      // because extractAttachments lifted inline non-text parts into the list.
      // Strict mode in searchMessages now demands disposition === "attachment".
      mockSearch.mockResolvedValue([1, 2]);
      mockFetch.mockReturnValue(
        (async function* () {
          // Inline banner — should NOT count as an attachment.
          yield {
            uid: 1,
            envelope: {},
            flags: [],
            bodyStructure: { disposition: "inline", type: "image/png", part: "2" },
          };
          // Real PDF attachment — should count.
          yield {
            uid: 2,
            envelope: {},
            flags: [],
            bodyStructure: { disposition: "attachment", type: "application/pdf", part: "2" },
          };
        })(),
      );

      const service = makeService();
      const result = await service.searchMessages("INBOX", { hasAttachment: true }, 10);
      expect(result.map((m) => m.uid)).toEqual([2]);
    });

    it("hasAttachment=true sets a default 5_000 byte floor on larger if not already set", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      await service.searchMessages("INBOX", { hasAttachment: true }, 10);
      expect(mockSearch).toHaveBeenCalledWith(expect.objectContaining({ larger: 5_000 }), { uid: true });
    });

    it("hasAttachment=true respects an explicit larger override", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      await service.searchMessages("INBOX", { hasAttachment: true, larger: 100_000 }, 10);
      expect(mockSearch).toHaveBeenCalledWith(expect.objectContaining({ larger: 100_000 }), { uid: true });
    });

    it("attachmentName filters by filename substring (case-insensitive) via bodyStructure", async () => {
      mockSearch.mockResolvedValue([1, 2]);
      // Message 1 has Invoice.pdf, Message 2 has photo.jpg.
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: {},
            flags: [],
            bodyStructure: {
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "Invoice-Q1.pdf" },
              type: "application/pdf",
              size: 30_000,
            },
          };
          yield {
            uid: 2,
            envelope: {},
            flags: [],
            bodyStructure: {
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "photo.jpg" },
              type: "image/jpeg",
              size: 60_000,
            },
          };
        })(),
      );

      const service = makeService();
      const result = await service.searchMessages("INBOX", { attachmentName: "invoice" }, 10);
      expect(result.map((m) => m.uid)).toEqual([1]);
    });

    it("attachmentType filters by MIME prefix via bodyStructure", async () => {
      mockSearch.mockResolvedValue([1, 2]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: {},
            flags: [],
            bodyStructure: {
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "doc.pdf" },
              type: "application/pdf",
              size: 30_000,
            },
          };
          yield {
            uid: 2,
            envelope: {},
            flags: [],
            bodyStructure: {
              part: "2",
              disposition: "attachment",
              dispositionParameters: { filename: "img.png" },
              type: "image/png",
              size: 30_000,
            },
          };
        })(),
      );

      const service = makeService();
      const result = await service.searchMessages("INBOX", { attachmentType: "application/pdf" }, 10);
      expect(result.map((m) => m.uid)).toEqual([1]);
    });

    it("attachment filters apply the 5 KB SIZE floor like hasAttachment", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      await service.searchMessages("INBOX", { attachmentName: "invoice" }, 10);
      expect(mockSearch).toHaveBeenCalledWith(expect.objectContaining({ larger: 5_000 }), { uid: true });
    });
  });

  describe("countMessages", () => {
    it("returns SEARCH result length without fetching envelopes", async () => {
      mockSearch.mockResolvedValue([1, 2, 3, 4, 5]);
      const service = makeService();
      const count = await service.countMessages("INBOX", { seen: false });
      expect(count).toBe(5);
      expect(mockFetch).not.toHaveBeenCalled();
    });

    it("returns 0 on empty SEARCH result", async () => {
      mockSearch.mockResolvedValue([]);
      const service = makeService();
      expect(await service.countMessages("INBOX", {})).toBe(0);
    });

    it("rejects a namespace container instead of returning 0", async () => {
      mockSearch.mockResolvedValue([]);
      mockList.mockResolvedValueOnce([
        {
          path: "Labels",
          name: "Labels",
          specialUse: "",
          flags: new Set(["\\NonExistent", "\\HasChildren"]),
          status: { messages: 0, unseen: 0 },
        },
      ]);
      const service = makeService();
      await expect(service.countMessages("Labels", {})).rejects.toThrow(
        /namespace container, not a selectable mailbox/,
      );
    });
  });

  describe("folderStats", () => {
    it("returns totals + scanned aggregations + truncation indicator", async () => {
      mockStatus.mockResolvedValue({ messages: 12000, unseen: 350 });
      mockSearch.mockResolvedValue(Array.from({ length: 12000 }, (_, i) => i + 1));
      mockFetch.mockReturnValue(
        (async function* () {
          yield { uid: 12000, envelope: { date: new Date("2026-05-20") }, size: 5000, flags: [] };
          yield { uid: 11999, envelope: { date: new Date("2026-05-19") }, size: 2000, flags: [] };
        })(),
      );

      const service = makeService();
      const stats = await service.folderStats("INBOX", 5000);

      expect(stats.total).toBe(12000);
      expect(stats.unread).toBe(350);
      expect(stats.scanLimit).toBe(5000);
      expect(stats.truncated).toBe(true);
      expect(stats.scanned).toBe(2); // mock only yielded 2 envelopes
      expect(stats.newest).toBe(new Date("2026-05-20").toISOString());
      expect(stats.oldest).toBe(new Date("2026-05-19").toISOString());
      expect(stats.totalBytes).toBe(7000);
    });

    it("rejects a namespace container instead of reporting empty stats", async () => {
      mockStatus.mockResolvedValueOnce({ messages: 0, unseen: 0 });
      mockSearch.mockResolvedValueOnce([]);
      mockList.mockResolvedValueOnce([
        {
          path: "Folders",
          name: "Folders",
          specialUse: "",
          flags: new Set(["\\Noselect", "\\HasChildren"]),
          status: { messages: 0, unseen: 0 },
        },
      ]);
      const service = makeService();
      await expect(service.folderStats("Folders", 5000)).rejects.toThrow(/namespace container/);
    });
  });

  describe("topSenders", () => {
    it("collapses 'Foo <foo@x.com>' and 'foo@x.com' into one bucket", async () => {
      mockSearch.mockResolvedValue([1, 2, 3]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: { from: [{ name: "Foo", address: "foo@x.com" }], date: new Date("2026-05-20") },
            flags: [],
          };
          yield { uid: 2, envelope: { from: [{ address: "foo@x.com" }], date: new Date("2026-05-22") }, flags: [] };
          yield { uid: 3, envelope: { from: [{ address: "BAR@y.com" }], date: new Date("2026-05-21") }, flags: [] };
        })(),
      );

      const service = makeService();
      const result = await service.topSenders("INBOX", {}, 10, 5000);
      const foo = result.rows.find((r) => r.from.toLowerCase().includes("foo@x.com"));
      expect(foo?.count).toBe(2);
      expect(foo?.lastDate).toBe(new Date("2026-05-22").toISOString());
    });

    it("uses the majority display name so one spoofed From can't poison the bucket label", async () => {
      // 3 messages from the SAME address: 2 with the legit name, 1 with a spoofed
      // display name. Majority wins — the spoofed text must NOT become the label.
      mockSearch.mockResolvedValue([1, 2, 3]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: { from: [{ name: "Acme Updates", address: "noreply@acme.com" }], date: new Date("2026-05-20") },
            flags: [],
          };
          yield {
            uid: 2,
            envelope: { from: [{ name: "Acme Updates", address: "noreply@acme.com" }], date: new Date("2026-05-21") },
            flags: [],
          };
          yield {
            uid: 3,
            envelope: {
              from: [{ name: "Security Team support@bank.com", address: "noreply@acme.com" }],
              date: new Date("2026-05-22"),
            },
            flags: [],
          };
        })(),
      );
      const service = makeService();
      const result = await service.topSenders("INBOX", {}, 10, 5000);
      const row = result.rows.find((r) => r.from.includes("noreply@acme.com"));
      expect(row?.count).toBe(3);
      expect(row?.from).toBe("Acme Updates <noreply@acme.com>");
      expect(row?.from).not.toContain("Security Team");
    });

    it("reports truncated:true when SEARCH result exceeds scanLimit", async () => {
      mockSearch.mockResolvedValue(Array.from({ length: 6000 }, (_, i) => i + 1));
      mockFetch.mockReturnValue((async function* () {})());

      const service = makeService();
      const result = await service.topSenders("INBOX", {}, 10, 5000);
      expect(result.truncated).toBe(true);
    });

    it("excludes the authenticated user's address from rows when excludeSelf=true", async () => {
      mockSearch.mockResolvedValue([1, 2, 3]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: { from: [{ address: "me@pm.me" }], date: new Date("2026-05-20") },
            flags: [],
          };
          yield {
            uid: 2,
            envelope: { from: [{ address: "me@pm.me" }], date: new Date("2026-05-21") },
            flags: [],
          };
          yield {
            uid: 3,
            envelope: { from: [{ address: "external@x.com" }], date: new Date("2026-05-22") },
            flags: [],
          };
        })(),
      );

      const service = makeService();
      const result = await service.topSenders("All Mail", {}, 10, 5000, {
        excludeSelf: true,
        userAddress: "me@pm.me",
      });
      expect(result.rows.map((r) => r.from)).toEqual(["external@x.com"]);
      expect(result.rows[0].direction).toBe("received");
    });

    it('tags rows as "self"/"received" when userAddress is provided without excludeSelf', async () => {
      mockSearch.mockResolvedValue([1, 2]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: { from: [{ address: "me@pm.me" }], date: new Date("2026-05-20") },
            flags: [],
          };
          yield {
            uid: 2,
            envelope: { from: [{ address: "external@x.com" }], date: new Date("2026-05-21") },
            flags: [],
          };
        })(),
      );

      const service = makeService();
      const result = await service.topSenders("All Mail", {}, 10, 5000, {
        userAddress: "me@pm.me",
      });
      const self = result.rows.find((r) => r.from === "me@pm.me");
      const ext = result.rows.find((r) => r.from === "external@x.com");
      expect(self?.direction).toBe("self");
      expect(ext?.direction).toBe("received");
    });
  });

  describe("listMessages snippets", () => {
    it("attaches a snippet when includeSnippet=true", async () => {
      mockSearch.mockResolvedValue([1]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield {
            uid: 1,
            envelope: { date: new Date("2026-05-20"), subject: "Hi", from: [{ address: "a@b.c" }], to: [] },
            flags: [],
            bodyStructure: { type: "text/plain", part: "1" },
          };
        })(),
      );
      mockFetchOne.mockResolvedValue({
        uid: 1,
        bodyStructure: { type: "text/plain", part: "1" },
      });
      mockDownload.mockResolvedValue({
        content: (async function* () {
          yield Buffer.from("Hello there  agent.  This is a long body with    extra whitespace runs.");
        })(),
        meta: {},
      });

      const service = makeService();
      const result = await service.listMessages("INBOX", 10, undefined, { includeSnippet: true });
      expect(result[0].snippet).toMatch(/Hello there agent/);
      expect(result[0].snippet?.length).toBeLessThanOrEqual(201); // 200 + ellipsis
    });

    it("returns empty snippet rather than failing the row when downloadPart throws", async () => {
      mockSearch.mockResolvedValue([1]);
      mockFetch.mockReturnValue(
        (async function* () {
          yield { uid: 1, envelope: {}, flags: [], bodyStructure: { type: "text/plain", part: "1" } };
        })(),
      );
      mockFetchOne.mockResolvedValue({
        uid: 1,
        bodyStructure: { type: "text/plain", part: "1" },
      });
      mockDownload.mockRejectedValue(new Error("download failed"));

      const service = makeService();
      const result = await service.listMessages("INBOX", 10, undefined, { includeSnippet: true });
      expect(result[0].snippet).toBe("");
    });
  });

  describe("moveThread", () => {
    it("acrossFolders=false acts only in the seed message's folder", async () => {
      // findByMessageId seed: returns UID 42 in INBOX
      // Thread walk: returns the seed (single-folder mode caps the folder list to [INBOX])
      mockSearch
        .mockResolvedValueOnce([42]) // findByMessageId in INBOX (firstFolderForMessageId)
        .mockResolvedValueOnce([42]) // getThreadByMessageId seed lookup in INBOX
        .mockResolvedValue([42]); // walkReferences within INBOX
      mockFetchOne.mockResolvedValue({
        uid: 42,
        envelope: { messageId: "<m@id>", date: new Date("2026-05-20") },
        headers: Buffer.from("References: <m@id>\r\n"),
      });
      // fetch is called twice in this flow: (1) thread envelope fetch returns full
      // envelope + flags; (2) bulkMove's existence pre-check only reads uid.
      // mockImplementation gives each call a fresh iterator (mockReturnValue would
      // hand out the same exhausted iterator on the second call).
      // Thread envelope fetch must include envelope.messageId so threadOp can
      // pick it up for per-folder re-resolution. bulkMove's existence pre-check
      // only reads uid, so the same iterator shape works for both calls.
      mockFetch.mockImplementation(() =>
        (async function* () {
          yield {
            uid: 42,
            envelope: { messageId: "<m@id>", date: new Date("2026-05-20"), from: [{ address: "a@b.c" }], to: [] },
            flags: [],
          };
        })(),
      );
      mockMessageMove.mockResolvedValue({ uidMap: new Map([[42, 442]]) });

      const service = makeService();
      const result = await service.moveThread("<m@id>", "Archive", false);

      expect(result.perFolder).toEqual([{ folder: "INBOX", affected: 1, notFound: [] }]);
      expect(result.total).toBe(1);
    });

    it("acrossFolders=true captures per-folder errors instead of aborting", async () => {
      // Thread spans INBOX (UID 42) + Sent (UID 7). The Sent bulkMove throws.
      // Expect: perFolder has both entries, the Sent entry carries `error`,
      // total counts only the successful INBOX move.
      // Stub at a higher level to bypass walkReferences mock-queue complexity.
      const service = makeService();

      vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([
        {
          uid: 42,
          folder: "INBOX",
          messageId: "<m@id>",
          subject: "x",
          from: "a@b.c",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
        },
        {
          uid: 7,
          folder: "Sent",
          messageId: "<m@id>",
          subject: "x",
          from: "a@b.c",
          to: "",
          date: "2026-05-21T00:00:00Z",
          flags: [],
        },
      ] as never);

      // threadOp now re-resolves UIDs per folder via resolveUidsByMessageIds, which
      // calls client.search under the lock. Return INBOX's UID first, Sent's second.
      mockSearch.mockResolvedValueOnce([42]).mockResolvedValueOnce([7]);

      vi.spyOn(service, "bulkMove").mockImplementation(async (folder) => {
        if (folder === "Sent") throw new Error("server hiccup");
        return { moved: 1, notFound: [], newUids: [442], destination: "Archive" };
      });

      const result = await service.moveThread("<m@id>", "Archive", true);

      expect(result.total).toBe(1); // only INBOX succeeded
      const inbox = result.perFolder.find((p) => p.folder === "INBOX");
      const sent = result.perFolder.find((p) => p.folder === "Sent");
      expect(inbox?.affected).toBe(1);
      expect(inbox?.error).toBeUndefined();
      expect(sent?.affected).toBe(0);
      expect(sent?.error).toMatch(/server hiccup/);
    });

    it("re-resolves UIDs per folder so stale-snapshot cascades report 0/0, not false notFound", async () => {
      // Regression guard for stale-snapshot cascade handling.
      //
      // Scenario: thread has members in INBOX and All Mail. previewThread snapshots
      // both. By the time we operate on All Mail, Proton has cascaded the INBOX→Trash
      // move and the All Mail UIDs have vanished. The old code reported them as
      // `notFound` (misleading — they did move); the new code re-resolves at mutation
      // time, finds an empty set, and reports a clean { affected: 0, notFound: [] }.
      const service = makeService();

      vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([
        {
          uid: 19,
          folder: "INBOX",
          messageId: "<a@id>",
          subject: "A",
          from: "x@y",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
        },
        {
          uid: 27,
          folder: "All Mail",
          messageId: "<a@id>",
          subject: "A",
          from: "x@y",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
        },
      ] as never);

      // First resolveUidsByMessageIds (INBOX) finds [19]; second (All Mail) finds []
      // because the prior INBOX→Trash op already stripped the All Mail label.
      mockSearch.mockResolvedValueOnce([19]).mockResolvedValueOnce([]);

      vi.spyOn(service, "bulkMove").mockResolvedValue({
        moved: 1,
        notFound: [],
        newUids: [99],
        destination: "Archive",
      });

      const result = await service.moveThread("<a@id>", "Archive", true);

      const inbox = result.perFolder.find((p) => p.folder === "INBOX");
      const allMail = result.perFolder.find((p) => p.folder === "All Mail");
      expect(inbox).toEqual({ folder: "INBOX", affected: 1, notFound: [] });
      // Critical assertion: All Mail reports clean 0/0, NOT { affected: 0, notFound: [27] }.
      expect(allMail).toEqual({ folder: "All Mail", affected: 0, notFound: [] });
      expect(result.total).toBe(1);
    });
  });

  // Regression guard for the v0.6.0 bug where thread dry-runs ignored acrossFolders.
  // The dry-run handlers in index.ts feed previewThread directly, so any change that
  // makes acrossFolders=false fan out broadly will break these.
  describe("previewThread", () => {
    it("acrossFolders=false scopes to the seed message's folder", async () => {
      const service = makeService();
      const spy = vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([]);
      // firstFolderForMessageId walks DEFAULT_THREAD_FOLDERS via findByMessageId,
      // which uses its own client + search. Stubbing the search result is enough:
      // INBOX hit → returns [INBOX] without probing Sent / All Mail.
      mockSearch.mockResolvedValueOnce([42]);

      await service.previewThread("<m@id>", false);

      expect(spy).toHaveBeenCalledWith("<m@id>", 1000, ["INBOX"]);
    });

    it("acrossFolders=true walks the default folder set", async () => {
      const service = makeService();
      const spy = vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([]);

      await service.previewThread("<m@id>", true);

      // undefined → getThreadByMessageId falls back to DEFAULT_THREAD_FOLDERS
      expect(spy).toHaveBeenCalledWith("<m@id>", 1000, undefined);
    });

    it("acrossFolders=true reroutes All Mail-only orphans to their real storage folder", async () => {
      // Scenario: a thread member lives in Folders/Archive. The default
      // INBOX+Sent+All Mail walk surfaces it only via All Mail (no INBOX/Sent
      // copy exists). Without rerouting, a thread mutation against this
      // entry would target "All Mail" and silently no-op under Proton's
      // label model.
      const service = makeService();
      vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([
        {
          uid: 120,
          folder: "All Mail",
          messageId: "<orphan@id>",
          subject: "S",
          from: "x@y",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
          otherFolders: [],
        },
      ] as never);

      // client.list() returns the candidate folder set. Folders/Archive
      // holds the message; the search there resolves to [42].
      mockList.mockResolvedValueOnce([
        { path: "INBOX", name: "INBOX", specialUse: "" },
        { path: "Sent", name: "Sent", specialUse: "\\Sent" },
        { path: "Folders/Archive", name: "Archive", specialUse: "" },
        { path: "All Mail", name: "All Mail", specialUse: "\\All" },
        { path: "Trash", name: "Trash", specialUse: "\\Trash" },
      ]);
      // INBOX and Sent return [], Folders/Archive returns [42].
      mockSearch.mockResolvedValueOnce([]).mockResolvedValueOnce([]).mockResolvedValueOnce([42]);

      const result = await service.previewThread("<orphan@id>", true);

      expect(result).toHaveLength(1);
      expect(result[0].folder).toBe("Folders/Archive");
      expect(result[0].messageId).toBe("<orphan@id>");
      // The UID must also be rewritten to the real folder's UID (42),
      // not left as the All Mail UID (120). The dry-run preview renders
      // `folder: UID`, so a stale UID would pair "Folders/Archive / UID 120"
      // where 120 is invalid in Archive — an agent copying that into a
      // single-message tool would hit the wrong message.
      expect(result[0].uid).toBe(42);
    });

    it("acrossFolders=true leaves non-orphan rows untouched", async () => {
      // A row with otherFolders set is already deduped — the dedupe found a
      // non-All-Mail copy. No rerouting needed.
      const service = makeService();
      vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([
        {
          uid: 19,
          folder: "INBOX",
          messageId: "<m@id>",
          subject: "S",
          from: "x@y",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
          otherFolders: ["All Mail"],
        },
      ] as never);

      const result = await service.previewThread("<m@id>", true);

      // No list / search calls fired — rerouteAllMailOrphans short-circuits.
      expect(mockList).not.toHaveBeenCalled();
      expect(result[0].folder).toBe("INBOX");
    });

    it("acrossFolders=false reroutes All Mail-only orphans too", async () => {
      // Regression guard: firstFolderForMessageId walks INBOX → Sent → All
      // Mail. For a message in Folders/Archive, the first two miss and the
      // walk returns ["All Mail"]. previewThread used to only run the
      // orphan-reroute for acrossFolders=true, so the acrossFolders=false
      // call left the row tagged folder="All Mail" — the live mutation then
      // no-op'd against the virtual All Mail mailbox under Proton's label
      // model. Same silent-no-op bug as the acrossFolders=true path above,
      // just on the other code path.
      const service = makeService();
      // firstFolderForMessageId search: INBOX miss, Sent miss, All Mail hit
      mockSearch.mockResolvedValueOnce([]).mockResolvedValueOnce([]).mockResolvedValueOnce([99]);
      vi.spyOn(service, "getThreadByMessageId").mockResolvedValue([
        {
          uid: 120,
          folder: "All Mail",
          messageId: "<archived@id>",
          subject: "S",
          from: "x@y",
          to: "",
          date: "2026-05-20T00:00:00Z",
          flags: [],
          otherFolders: [],
        },
      ] as never);
      // rerouteAllMailOrphans walks user folders looking for the Message-ID
      mockList.mockResolvedValueOnce([
        { path: "INBOX", name: "INBOX", specialUse: "" },
        { path: "Sent", name: "Sent", specialUse: "\\Sent" },
        { path: "Folders/Archive", name: "Archive", specialUse: "" },
        { path: "All Mail", name: "All Mail", specialUse: "\\All" },
      ]);
      // INBOX miss, Sent miss, Folders/Archive hit
      mockSearch.mockResolvedValueOnce([]).mockResolvedValueOnce([]).mockResolvedValueOnce([42]);

      const result = await service.previewThread("<archived@id>", false);

      expect(result[0].folder).toBe("Folders/Archive");
      // UID rewritten to Archive's UID (42), not the All Mail UID (120).
      expect(result[0].uid).toBe(42);
    });
  });
});
