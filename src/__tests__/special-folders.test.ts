import { describe, it, expect, vi } from "vitest";
import { SpecialFolderResolver } from "../special-folders.js";

function makeMockClient(listResult: Array<{ path: string; specialUse?: string }>) {
  return {
    list: vi.fn().mockResolvedValue(listResult),
  } as unknown as import("imapflow").ImapFlow;
}

describe("SpecialFolderResolver", () => {
  it("resolves trash/junk/archive/sent/drafts from specialUse annotations", async () => {
    const resolver = new SpecialFolderResolver();
    const client = makeMockClient([
      { path: "INBOX" },
      { path: "Trash", specialUse: "\\Trash" },
      { path: "Spam", specialUse: "\\Junk" },
      { path: "Archive", specialUse: "\\Archive" },
      { path: "Sent", specialUse: "\\Sent" },
      { path: "Drafts", specialUse: "\\Drafts" },
    ]);

    const result = await resolver.resolve(client);

    expect(result).toEqual({
      trash: "Trash",
      junk: "Spam",
      archive: "Archive",
      sent: "Sent",
      drafts: "Drafts",
    });
  });

  it("caches the result after the first resolve", async () => {
    const resolver = new SpecialFolderResolver();
    const client = makeMockClient([{ path: "Trash", specialUse: "\\Trash" }]);

    await resolver.resolve(client);
    await resolver.resolve(client);
    await resolver.resolve(client);

    expect(client.list).toHaveBeenCalledTimes(1);
  });

  it("shares one in-flight promise across concurrent first-callers", async () => {
    const resolver = new SpecialFolderResolver();
    let resolveList!: (v: Array<{ path: string; specialUse?: string }>) => void;
    const client = {
      list: vi.fn().mockReturnValue(
        new Promise((res) => {
          resolveList = res;
        }),
      ),
    } as unknown as import("imapflow").ImapFlow;

    const p1 = resolver.resolve(client);
    const p2 = resolver.resolve(client);
    const p3 = resolver.resolve(client);

    resolveList([{ path: "Trash", specialUse: "\\Trash" }]);

    await Promise.all([p1, p2, p3]);
    expect(client.list).toHaveBeenCalledTimes(1);
  });

  it("returns undefined for slots with no specialUse annotation", async () => {
    const resolver = new SpecialFolderResolver();
    const client = makeMockClient([{ path: "INBOX" }, { path: "Random" }]);

    const result = await resolver.resolve(client);

    expect(result.trash).toBeUndefined();
    expect(result.junk).toBeUndefined();
    expect(result.archive).toBeUndefined();
  });

  it("invalidate() forces re-resolution on next call", async () => {
    const resolver = new SpecialFolderResolver();
    const client = makeMockClient([{ path: "Trash", specialUse: "\\Trash" }]);

    await resolver.resolve(client);
    resolver.invalidate();
    await resolver.resolve(client);

    expect(client.list).toHaveBeenCalledTimes(2);
  });
});
