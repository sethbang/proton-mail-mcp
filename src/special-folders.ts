import type { ImapFlow } from "imapflow";

export interface SpecialFolders {
  trash?: string;
  junk?: string;
  archive?: string;
  sent?: string;
  drafts?: string;
}

/**
 * Resolves IMAP special-use folder annotations (\Trash, \Junk, \Archive, \Sent, \Drafts)
 * to their actual mailbox paths. Caches the result so it costs one LIST per server lifetime.
 *
 * Concurrent first-callers share a single in-flight promise. Falls back to undefined when
 * a slot has no special-use annotation; callers should default to the literal name in that case.
 */
export class SpecialFolderResolver {
  private cache?: SpecialFolders;
  private inFlight?: Promise<SpecialFolders>;

  async resolve(client: ImapFlow): Promise<SpecialFolders> {
    if (this.cache) return this.cache;
    if (this.inFlight) return this.inFlight;

    this.inFlight = (async () => {
      const mailboxes = await client.list();
      const out: SpecialFolders = {};
      for (const mb of mailboxes) {
        switch (mb.specialUse) {
          case "\\Trash":
            out.trash = mb.path;
            break;
          case "\\Junk":
            out.junk = mb.path;
            break;
          case "\\Archive":
            out.archive = mb.path;
            break;
          case "\\Sent":
            out.sent = mb.path;
            break;
          case "\\Drafts":
            out.drafts = mb.path;
            break;
        }
      }
      this.cache = out;
      this.inFlight = undefined;
      return out;
    })();

    return this.inFlight;
  }

  invalidate(): void {
    this.cache = undefined;
    this.inFlight = undefined;
  }
}
