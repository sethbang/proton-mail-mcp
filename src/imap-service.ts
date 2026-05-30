import { ImapFlow } from "imapflow";
import type { FetchMessageObject, MessageAddressObject, MessageStructureObject, SearchObject } from "imapflow";
import { convert } from "html-to-text";
import { validateFolderPath, validatePartNumber, assertSinceBeforeNotIdentical } from "./validation.js";
import { SpecialFolderResolver, type SpecialFolders } from "./special-folders.js";

export type { SpecialFolders } from "./special-folders.js";

export interface ImapConfig {
  host: string;
  port: number;
  secure: boolean;
  auth: {
    user: string;
    pass: string;
  };
  debug?: boolean;
  connectionTimeout?: number;
  maxAttachmentSize?: number;
}

export interface MessageSummary {
  uid: number;
  subject: string;
  from: string;
  to: string;
  date: string;
  flags: string[];
  snippet?: string;
  /**
   * Message-ID header value. Populated by every fetch path that requests envelopes.
   * Empty string when the envelope had no Message-ID (rare in practice). Used by
   * threadOp to re-resolve per-folder UIDs at mutation time, sidestepping the
   * stale-snapshot problem on Proton's label-cascade model.
   */
  messageId?: string;
}

export interface AttachmentMeta {
  partNumber: string;
  filename: string;
  contentType: string;
  /**
   * Estimated decoded (real file) size in bytes — roughly what a user sees when
   * they save the attachment. IMAP reports the *encoded* octet count; for base64
   * parts that's ~37% larger than the file, so we convert (see
   * `decodedSizeEstimate`). **Approximate (±~2 bytes)** — the true decoded size
   * isn't recoverable from the encoded count without downloading the part, so do
   * NOT treat this as byte-exact (don't pre-allocate buffers or assert on it).
   * Displayed with a leading `~`. The exact saved byte count is returned by
   * `download_attachment` (it reports `bytes.length` of the decoded payload).
   */
  size: number;
  /**
   * Raw IMAP-reported octet count = the size of the part as it travels on the
   * wire (base64-encoded for binary attachments). This is what an inline
   * base64 response actually costs, so the inline-size cap compares against it.
   */
  encodedSize: number;
}

export interface MessageDetail {
  uid: number;
  subject: string;
  from: string;
  to: string;
  cc: string;
  date: string;
  messageId: string;
  flags: string[];
  body: string;
  bodyFormat: "text" | "html-stripped" | "html";
  truncated: boolean;
  originalLength?: number;
  attachments: AttachmentMeta[];
  /** Additional headers parsed from the raw message, only when `showHeaders: true` was passed. */
  extraHeaders?: Record<string, string>;
}

/** Default folder set walked by get_thread when searching across a mailbox. */
export const DEFAULT_THREAD_FOLDERS = ["INBOX", "Sent", "All Mail"];

/** Extra headers requested/parsed when readMessage is called with `showHeaders: true`. */
const EXTRA_HEADER_NAMES = ["in-reply-to", "references", "reply-to", "list-unsubscribe", "list-id"];

export interface MoveResult {
  success: boolean;
  newUid?: number;
  destination: string;
}

export interface BulkMoveResult {
  moved: number;
  notFound: number[];
  newUids: number[];
  destination: string;
}

export interface BulkDeleteResult {
  deleted: number;
  notFound: number[];
  newUids?: number[]; // only when permanent=false (move to Trash supports UIDPLUS)
  destination?: string; // resolved Trash path when permanent=false
}

export interface BulkFlagResult {
  affected: number;
  notFound: number[];
  /** Flags whose STORE the server silently dropped on every affected UID. */
  notApplied: string[];
}

export interface BulkLabelResult {
  affected: number;
  notFound: number[];
  /**
   * Labels whose add or remove had no effect on any of the affected UIDs:
   * - add: every COPY was a no-op (label already applied to every UID)
   * - remove: no label-mailbox entries matched any of the affected UIDs
   */
  notApplied: string[];
}

export interface SearchCriteria {
  from?: string;
  to?: string;
  subject?: string;
  body?: string;
  since?: string;
  before?: string;
  seen?: boolean;
  flagged?: boolean;
  larger?: number;
  smaller?: number;
  listId?: string;
  hasAttachment?: boolean;
  /** Case-insensitive substring filter on attachment filenames (e.g. "invoice", ".pdf"). */
  attachmentName?: string;
  /** Case-insensitive MIME-type prefix filter on attachments (e.g. "application/pdf", "image/"). */
  attachmentType?: string;
}

export interface FolderInfo {
  path: string;
  name: string;
  specialUse: string;
  messages: number;
  unseen: number;
  /**
   * True when the IMAP server flagged the mailbox as `\Noselect` — typically a
   * namespace container (e.g. Proton's top-level "Folders" / "Labels") that
   * holds nested mailboxes but cannot itself be opened, written to, or used as
   * a move destination. These entries appear alongside real folders with 0
   * messages and are easily mistaken for empty mailboxes one could plant
   * messages into.
   */
  noSelect?: boolean;
}

export interface FolderStats {
  folder: string;
  total: number;
  unread: number;
  scanned: number;
  truncated: boolean;
  scanLimit: number;
  oldest?: string;
  newest?: string;
  totalBytes?: number;
}

export interface TopSenderRow {
  from: string;
  count: number;
  lastDate?: string;
  /**
   * Set when the caller passes `userAddress` to `topSenders`. "self" when the
   * bucket's address matches the authenticated user (typically rows that come
   * from a folder spanning Sent, like All Mail); "received" otherwise.
   */
  direction?: "self" | "received";
}

export interface TopSendersResult {
  rows: TopSenderRow[];
  scanned: number;
  truncated: boolean;
  scanLimit: number;
}

export interface PerFolderResult {
  folder: string;
  affected: number;
  notFound: number[];
  error?: string;
}

export interface ThreadOpResult {
  perFolder: PerFolderResult[];
  total: number;
}

function formatAddress(addr?: MessageAddressObject[]): string {
  if (!addr || addr.length === 0) return "";
  return addr.map((a) => (a.name ? `${a.name} <${a.address}>` : a.address || "")).join(", ");
}

/** Default maximum body length in characters returned by readMessage. */
const DEFAULT_MAX_BODY_LENGTH = 50_000;

/**
 * Find the MIME part number for a given content type in a bodyStructure tree.
 * Returns the first match (depth-first). Parts marked `Content-Disposition: attachment`
 * are skipped — selecting a text/plain attachment as the body would silently return
 * the attachment's content to the caller instead of the real body.
 *
 * Single-part messages have no `part` field on the root node — IMAP
 * implicitly treats those as part "1".
 */
function findPartNumber(structure: MessageStructureObject, targetType: string): string | undefined {
  if (structure.disposition === "attachment") return undefined;
  if (structure.type === targetType) return structure.part || "1";
  if (structure.childNodes) {
    for (const child of structure.childNodes) {
      const found = findPartNumber(child, targetType);
      if (found) return found;
    }
  }
  return undefined;
}

/**
 * Convert HTML to readable plain text using html-to-text.
 * Preserves links as `text [url]`, skips images/tracking pixels,
 * and handles whitespace, entities, and list formatting.
 *
 * `stripUrls: true` drops anchor hrefs entirely — only the visible link text
 * remains. Useful for summarizing newsletters without burning tokens on
 * tracking URLs.
 */
function stripHtml(html: string, stripUrls = false): string {
  const anchorOptions = stripUrls ? { ignoreHref: true } : { hideLinkHrefIfSameAsText: true };
  return convert(html, {
    wordwrap: false,
    selectors: [
      { selector: "a", options: anchorOptions },
      { selector: "img", format: "skip" },
    ],
  });
}

/**
 * Parse RFC 5322 header fields from a raw headers buffer.
 * Handles folded continuation lines (leading WSP on subsequent lines).
 * Returns a lowercase-keyed map of header name → value.
 */
function parseHeaders(raw: string, wanted: readonly string[]): Record<string, string> {
  const wantedSet = new Set(wanted.map((n) => n.toLowerCase()));
  // Unfold: join continuation lines (a line starting with space/tab belongs to the previous header).
  const unfolded = raw.replace(/\r?\n[ \t]+/g, " ");
  const out: Record<string, string> = {};
  for (const line of unfolded.split(/\r?\n/)) {
    const match = line.match(/^([^\s:]+):\s*(.*)$/);
    if (!match) continue;
    const name = match[1].toLowerCase();
    if (!wantedSet.has(name)) continue;
    out[name] = match[2].trim();
  }
  return out;
}

/**
 * Recognize imapflow / IMAP server responses that indicate the target mailbox doesn't exist.
 *
 * imapflow sets canonical signals on the error we should trust first:
 *  - `err.mailboxMissing = true` — set by getMailboxLock when SELECT fails NO and a LIST probe
 *    confirms the mailbox doesn't exist (imap-flow.js ~3580).
 *  - `err.serverResponseCode` — e.g. "NONEXISTENT", "TRYCREATE" from the NO response code.
 *
 * The text fallback covers cases where imapflow hasn't annotated the error (defense in depth).
 */
function isMailboxMissingError(err: unknown): boolean {
  if (!(err instanceof Error)) return false;
  const e = err as Error & { mailboxMissing?: boolean; serverResponseCode?: string };
  if (e.mailboxMissing === true) return true;
  if (e.serverResponseCode === "NONEXISTENT" || e.serverResponseCode === "TRYCREATE") return true;
  const msg = err.message.toLowerCase();
  return (
    msg.includes("no such mailbox") || msg.includes("mailbox doesn't exist") || msg.includes("mailbox does not exist")
  );
}

/**
 * IMAP service for reading emails via Proton Mail Bridge.
 * Each operation creates a fresh connection to avoid stale/idle connection issues.
 */
export class ImapService {
  private config: ImapConfig;
  private debug: boolean;
  private maxAttachmentSize: number;
  private specialFolderResolver = new SpecialFolderResolver();

  constructor(config: ImapConfig) {
    this.config = config;
    this.debug = config.debug || false;
    this.maxAttachmentSize = config.maxAttachmentSize ?? 25 * 1024 * 1024;
  }

  private createClient(): ImapFlow {
    return new ImapFlow({
      host: this.config.host,
      port: this.config.port,
      secure: this.config.secure,
      auth: this.config.auth,
      logger: false,
      connectionTimeout: this.config.connectionTimeout ?? 30_000,
    });
  }

  private log(message: string): void {
    if (this.debug) {
      console.error(message);
    }
  }

  /**
   * Acquire a mailbox lock, translating imapflow's mailbox-missing signals into a
   * uniform "Folder not found: <path>" error. Every folder-taking operation should
   * route through this instead of calling getMailboxLock directly.
   */
  private async lockFolder(client: ImapFlow, folder: string): ReturnType<ImapFlow["getMailboxLock"]> {
    try {
      return await client.getMailboxLock(folder);
    } catch (err) {
      if (isMailboxMissingError(err)) {
        throw new Error(`Folder not found: ${folder}`);
      }
      throw err;
    }
  }

  /**
   * Decide whether a LIST entry is a non-selectable namespace container (e.g.
   * Proton's top-level "Folders" / "Labels" nodes) rather than a real mailbox.
   *
   * The canonical IMAP signal is the `\Noselect` / `\NonExistent` flag
   * (RFC 3501), but three Proton Mail Bridge quirks keep the flag from being
   * authoritative on its own, so it is overridden by positive evidence of
   * selectability:
   *
   *   1. `listed === false` is NOT a selectability signal — it means "appeared
   *      in LSUB but not LIST" (subscription state). Proton surfaces selectable
   *      labels via LSUB with `listed: false`, so that branch is gated on an
   *      *absent* status object — it can only catch a genuine container that
   *      also reports no live counts.
   *   2. The bridge has been observed returning a populated label with BOTH a
   *      `\Noselect`-family flag AND inline STATUS counts (via LIST-STATUS). A
   *      mailbox reporting a live message count is selectable by definition —
   *      you cannot hold messages in something you can't open — so a positive
   *      count overrides the flag.
   *   3. Even at zero messages, a leaf mailbox (`\HasNoChildren`, no nested
   *      mailboxes) that returned a STATUS object is selectable — a STATUS
   *      succeeds only on an openable mailbox. A genuine container carries
   *      `\HasChildren`. So a successful STATUS on a childless mailbox also
   *      overrides a spurious `\Noselect` flag, rescuing an *empty* label the
   *      bridge mis-flags.
   */
  private static computeNoSelect(mb: {
    flags?: Set<string>;
    listed?: boolean;
    status?: { messages?: number };
  }): boolean {
    const flags = mb.flags instanceof Set ? mb.flags : new Set<string>();
    const flagNoSelect = flags.has("\\Noselect") || flags.has("\\NonExistent");
    const hasNoChildren = flags.has("\\HasNoChildren");
    const hasLiveStatus = mb.status !== undefined;
    const hasMessages = (mb.status?.messages ?? 0) > 0;
    // Positive evidence of selectability: a live message count, or a STATUS that
    // succeeded on a mailbox the server *explicitly* marks childless
    // (`\HasNoChildren`). Requiring the explicit flag — rather than the mere
    // absence of `\HasChildren` — keeps genuine containers tagged even on a
    // server that doesn't advertise the CHILDREN extension (RFC 3348) at all:
    // there, neither children flag is present, so the mailbox falls through to
    // the `\Noselect` check instead of being mistaken for a selectable leaf.
    const looksSelectable = hasMessages || (hasNoChildren && hasLiveStatus);
    return !looksSelectable && (flagNoSelect || (mb.listed === false && !hasLiveStatus));
  }

  /**
   * List available mailbox folders with message counts.
   */
  async listFolders(): Promise<FolderInfo[]> {
    this.log("[IMAP] Listing folders");
    const client = this.createClient();
    try {
      await client.connect();
      const mailboxes = await client.list({
        statusQuery: { messages: true, unseen: true },
      });
      return mailboxes.map((mb) => ({
        path: mb.path,
        name: mb.name,
        specialUse: mb.specialUse || "",
        messages: mb.status?.messages ?? 0,
        unseen: mb.status?.unseen ?? 0,
        noSelect: ImapService.computeNoSelect(mb),
      }));
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Throw an actionable error if `folder` is a non-selectable namespace
   * container (Proton's top-level "Folders" / "Labels" nodes).
   *
   * Proton Mail Bridge lets SELECT succeed on these containers and reports them
   * empty, so a read that returns zero results can't tell "genuinely empty
   * mailbox" from "not a mailbox at all" — `list_folders` annotates them, but
   * `list_messages` / `search_messages` / `count_messages` / `folder_stats`
   * would otherwise present them as ordinary empty folders. Callers invoke this
   * ONLY on the zero-result path, so the extra LIST never touches the common
   * case. Uses a fresh short-lived connection to avoid interacting with any
   * mailbox lock the caller is holding. Reuses the same non-selectable
   * determination as `listFolders` via `computeNoSelect`.
   */
  private async assertSelectableFolder(folder: string): Promise<void> {
    let isContainer = false;
    const client = this.createClient();
    try {
      await client.connect();
      const mailboxes = await client.list({ statusQuery: { messages: true, unseen: true } });
      const entry = Array.isArray(mailboxes) ? mailboxes.find((mb) => mb.path === folder) : undefined;
      isContainer = Boolean(entry && ImapService.computeNoSelect(entry));
    } catch {
      // If selectability can't be determined (e.g. LIST failed), don't mask the
      // caller's original empty result with an unrelated error — skip the guard.
      isContainer = false;
    } finally {
      await client.logout().catch(() => {});
    }
    // Throw outside the try so the catch above can't swallow our own signal.
    if (isContainer) {
      throw new Error(
        `"${folder}" is a namespace container, not a selectable mailbox — it holds nested mailboxes but can't store messages. Pick a child mailbox (e.g. "${folder}/<name>"); run list_folders to see what's under it.`,
      );
    }
  }

  /**
   * Resolve special-use folder paths (\Trash, \Junk, \Archive, \Sent, \Drafts) via LIST.
   * Cached after the first call. Opens its own connection.
   */
  async getSpecialFolders(): Promise<SpecialFolders> {
    this.log("[IMAP] Resolving special-use folders");
    const client = this.createClient();
    try {
      await client.connect();
      return await this.specialFolderResolver.resolve(client);
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Fetch envelopes for a set of UIDs, sort by date descending, and return
   * the top `limit` messages. When the candidate set exceeds MAX_ENVELOPE_FETCH
   * we fall back to taking the highest UIDs as a heuristic.
   */
  private async fetchSortAndLimit(client: ImapFlow, uids: number[], limit: number): Promise<MessageSummary[]> {
    if (uids.length === 0) return [];

    // For up to 500 UIDs, fetch all envelopes so date sort is exact.
    // Beyond that, take the highest UIDs as an approximation.
    const MAX_ENVELOPE_FETCH = 500;
    const selectedUids = uids.length <= MAX_ENVELOPE_FETCH ? uids : uids.slice(-MAX_ENVELOPE_FETCH);
    const fetchRange = selectedUids.join(",");

    const messages: MessageSummary[] = [];
    for await (const msg of client.fetch(fetchRange, { uid: true, envelope: true, flags: true }, { uid: true })) {
      messages.push(this.toSummary(msg));
    }

    // Sort newest first by date, falling back to UID descending
    messages.sort((a, b) => {
      const da = a.date ? new Date(a.date).getTime() : 0;
      const db = b.date ? new Date(b.date).getTime() : 0;
      return da !== db ? db - da : b.uid - a.uid;
    });
    return messages.slice(0, limit);
  }

  /**
   * Fetch envelopes for the highest `limit` UIDs in `uids` and return them
   * sorted by UID descending (arrival order, newest first).
   *
   * Unlike `fetchSortAndLimit` (date-sorted, with a 500-UID candidate cap),
   * this selects the top `limit` UIDs *before* fetching, so it touches at most
   * `limit` envelopes and the result is a contiguous UID window. That makes
   * `beforeUid` pagination exact — each page is a clean UID range with no
   * skips or duplicates, even in folders where UID order disagrees with date
   * order (All Mail, or any folder holding moved messages). `search()` returns
   * UIDs in ascending order, so the tail is the highest set.
   */
  private async fetchUidsDescending(client: ImapFlow, uids: number[], limit: number): Promise<MessageSummary[]> {
    if (uids.length === 0) return [];
    const topUids = uids.slice(-limit);
    const messages: MessageSummary[] = [];
    for await (const msg of client.fetch(
      topUids.join(","),
      { uid: true, envelope: true, flags: true },
      { uid: true },
    )) {
      messages.push(this.toSummary(msg));
    }
    messages.sort((a, b) => b.uid - a.uid);
    return messages;
  }

  async listMessages(
    folder: string,
    limit: number,
    beforeUid?: number,
    options?: { includeSnippet?: boolean; sortByUid?: boolean },
  ): Promise<MessageSummary[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Listing messages from ${folder}, limit ${limit}${beforeUid ? `, beforeUid ${beforeUid}` : ""}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        let result: number[] | false;
        if (beforeUid) {
          result = await client.search({ uid: `1:${beforeUid - 1}` }, { uid: true });
        } else {
          result = await client.search({ all: true }, { uid: true });
        }
        if (!result || result.length === 0) {
          await this.assertSelectableFolder(folder);
          return [];
        }

        // sortByUid: exact, skip-free pagination in UID (arrival) order.
        // Default: date-sorted "newest first" — convenient, but its beforeUid
        // pagination is approximate when UID order ≠ date order (see helper docs).
        const messages = options?.sortByUid
          ? await this.fetchUidsDescending(client, result, limit)
          : await this.fetchSortAndLimit(client, result, limit);
        if (options?.includeSnippet) {
          await this.attachSnippets(client, messages);
        }
        return messages;
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Read a single message by UID, including body content.
   *
   * Uses `client.download()` to fetch body parts, which automatically decodes
   * quoted-printable and base64 transfer encodings.
   *
   * Prefers text/plain when available. Falls back to text/html with tags
   * stripped. Pass `preferHtml: true` to get raw HTML instead.
   *
   * Body is truncated to `maxBodyLength` characters (default 50 000).
   */
  async readMessage(
    folder: string,
    uid: number,
    options?: { preferHtml?: boolean; maxBodyLength?: number; showHeaders?: boolean; stripUrls?: boolean },
  ): Promise<MessageDetail> {
    validateFolderPath(folder);
    const preferHtml = options?.preferHtml ?? false;
    const maxLen = options?.maxBodyLength ?? DEFAULT_MAX_BODY_LENGTH;
    const showHeaders = options?.showHeaders ?? false;
    const stripUrls = options?.stripUrls ?? false;
    this.log(`[IMAP] Reading message UID ${uid} from ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        // Step 1: fetch metadata and structure (no body content yet)
        const fetchOptions: Parameters<typeof client.fetchOne>[1] = {
          uid: true,
          envelope: true,
          flags: true,
          bodyStructure: true,
        };
        if (showHeaders) {
          fetchOptions.headers = [...EXTRA_HEADER_NAMES];
        }
        const msg = await client.fetchOne(String(uid), fetchOptions, { uid: true });

        if (!msg) {
          throw new Error(`Message UID ${uid} not found in ${folder}`);
        }

        const attachments = msg.bodyStructure ? ImapService.extractAttachments(msg.bodyStructure) : [];

        // Step 2: determine which part to download
        let body = "";
        let bodyFormat: "text" | "html-stripped" | "html" = "text";

        const textPartNum = msg.bodyStructure ? findPartNumber(msg.bodyStructure, "text/plain") : undefined;
        const htmlPartNum = msg.bodyStructure ? findPartNumber(msg.bodyStructure, "text/html") : undefined;

        if (preferHtml && htmlPartNum) {
          body = await this.downloadPart(client, uid, htmlPartNum);
          bodyFormat = "html";
        } else if (textPartNum) {
          body = await this.downloadPart(client, uid, textPartNum);
          bodyFormat = "text";
        } else if (htmlPartNum) {
          const rawHtml = await this.downloadPart(client, uid, htmlPartNum);
          body = stripHtml(rawHtml, stripUrls);
          bodyFormat = "html-stripped";
        } else if (msg.bodyStructure) {
          // Fallback: no text/plain or text/html found (e.g. PGP-encrypted).
          // Download part "1" as raw text so we return *something* rather than empty.
          try {
            body = await this.downloadPart(client, uid, "1");
            bodyFormat = "text";
          } catch {
            // Part "1" may not exist or may not be downloadable — leave body empty
          }
        }

        // Step 3: truncate if needed
        let truncated = false;
        let originalLength: number | undefined;
        if (body.length > maxLen) {
          originalLength = body.length;
          body = body.slice(0, maxLen);
          truncated = true;
        }

        let extraHeaders: Record<string, string> | undefined;
        if (showHeaders && msg.headers) {
          const raw = Buffer.isBuffer(msg.headers) ? msg.headers.toString("utf-8") : String(msg.headers);
          extraHeaders = parseHeaders(raw, EXTRA_HEADER_NAMES);
        }

        return {
          uid: msg.uid,
          subject: msg.envelope?.subject || "",
          from: formatAddress(msg.envelope?.from),
          to: formatAddress(msg.envelope?.to),
          cc: formatAddress(msg.envelope?.cc),
          date: msg.envelope?.date?.toISOString() || "",
          messageId: msg.envelope?.messageId || "",
          flags: msg.flags ? [...msg.flags] : [],
          body,
          bodyFormat,
          truncated,
          originalLength,
          attachments,
          extraHeaders,
        };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Download a single MIME part as a decoded UTF-8 string.
   */
  private async downloadPart(client: ImapFlow, uid: number, partNumber: string): Promise<string> {
    const { content } = await client.download(String(uid), partNumber, { uid: true });
    const chunks: Buffer[] = [];
    for await (const chunk of content) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    return Buffer.concat(chunks).toString("utf-8");
  }

  /**
   * Build an imapflow SearchObject from a SearchCriteria object.
   * Only the `from`/`to`/`subject`/`body`/`since`/`before`/`seen`/`flagged`
   * subset is used here; larger/smaller/listId are added for bulk operations.
   * hasAttachment is NOT applied at this layer (requires post-filter on bodyStructure).
   */
  private buildSearchQuery(criteria: SearchCriteria): SearchObject {
    // Same-day since/before is silently empty under IMAP's exclusive BEFORE.
    // Throw at the query-build boundary so every caller (search, count, bulk,
    // top_senders) surfaces it the same way.
    assertSinceBeforeNotIdentical(criteria.since, criteria.before);
    const q: SearchObject = {};
    if (criteria.from) q.from = criteria.from;
    if (criteria.to) q.to = criteria.to;
    if (criteria.subject) q.subject = criteria.subject;
    if (criteria.body) q.body = criteria.body;
    if (criteria.since) q.since = criteria.since;
    if (criteria.before) q.before = criteria.before;
    if (criteria.seen !== undefined) q.seen = criteria.seen;
    if (criteria.flagged !== undefined) q.flagged = criteria.flagged;
    if (criteria.larger !== undefined) q.larger = criteria.larger;
    if (criteria.smaller !== undefined) q.smaller = criteria.smaller;
    if (criteria.listId) {
      q.header = { ...(q.header || {}), "List-Id": criteria.listId };
    }
    return q;
  }

  /**
   * Search messages in a folder by criteria.
   */
  async searchMessages(
    folder: string,
    criteria: SearchCriteria,
    limit: number,
    options?: { includeSnippet?: boolean },
  ): Promise<MessageSummary[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Searching in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        // hasAttachment, attachmentName, and attachmentType all require a
        // bodyStructure post-filter. Coalesce them into one predicate so the
        // SEARCH+FETCH+filter pipeline runs once.
        const needsAttachmentScan = Boolean(
          criteria.hasAttachment || criteria.attachmentName || criteria.attachmentType,
        );
        const query: SearchObject = this.buildSearchQuery({
          ...criteria,
          // None of the attachment predicates are native IMAP — drop them from the SEARCH.
          hasAttachment: undefined,
          attachmentName: undefined,
          attachmentType: undefined,
          // Apply a 5 KB size floor when filtering by attachment; it's a heuristic
          // that drops messages too small to plausibly carry one. Previously
          // 25 KB, which filtered out legitimate small attachments (text/JSON/
          // tiny images) — particularly in Sent, where Proton stores
          // compact copies that fit under the earlier threshold. The 500-
          // candidate cap downstream still bounds worst-case bodyStructure fetch.
          larger: needsAttachmentScan && criteria.larger === undefined ? 5_000 : criteria.larger,
        });

        const result = await client.search(query, { uid: true });
        if (!result || result.length === 0) {
          await this.assertSelectableFolder(folder);
          return [];
        }

        if (needsAttachmentScan) {
          const nameNeedle = criteria.attachmentName?.toLowerCase();
          const typeNeedle = criteria.attachmentType?.toLowerCase();
          // For attachment filters, do a single combined fetch (envelope + flags + bodyStructure)
          // so we can post-filter without a second round trip. Cap candidates at 500.
          const candidates = result.length <= 500 ? result : result.slice(-500);
          const summaries: MessageSummary[] = [];
          const matched = new Set<number>();
          for await (const msg of client.fetch(
            candidates.join(","),
            { uid: true, envelope: true, flags: true, bodyStructure: true },
            { uid: true },
          )) {
            summaries.push(this.toSummary(msg));
            if (!msg.bodyStructure) continue;
            // strict: true → only Content-Disposition: attachment parts count,
            // not inline-rendered images. Newsletters with banner images would
            // otherwise be reported as "has attachments" in the user-intuitive
            // sense.
            const attachments = ImapService.extractAttachments(msg.bodyStructure as MessageStructureObject, {
              strict: true,
            });
            if (attachments.length === 0) continue;
            const passes = attachments.some(
              (a) =>
                (!nameNeedle || a.filename.toLowerCase().includes(nameNeedle)) &&
                (!typeNeedle || a.contentType.toLowerCase().startsWith(typeNeedle)),
            );
            if (passes) matched.add(msg.uid);
          }
          summaries.sort((a, b) => {
            const da = a.date ? new Date(a.date).getTime() : 0;
            const db = b.date ? new Date(b.date).getTime() : 0;
            return da !== db ? db - da : b.uid - a.uid;
          });
          const filtered = summaries.filter((m) => matched.has(m.uid)).slice(0, limit);
          if (options?.includeSnippet) {
            await this.attachSnippets(client, filtered);
          }
          return filtered;
        }

        const messages = await this.fetchSortAndLimit(client, result, limit);
        if (options?.includeSnippet) {
          await this.attachSnippets(client, messages);
        }
        return messages;
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Resolve UIDs from a SearchCriteria object by running an IMAP SEARCH.
   * Used as the `match` path for bulk operations.
   * Rejects `hasAttachment` since that filter requires fetching bodyStructure.
   */
  async resolveUidsFromCriteria(folder: string, criteria: SearchCriteria): Promise<number[]> {
    if (criteria.hasAttachment !== undefined || criteria.attachmentName || criteria.attachmentType) {
      throw new Error(
        "Attachment-based filters (`hasAttachment`, `attachmentName`, `attachmentType`) are only supported by search_messages, not bulk match criteria. " +
          "Use search_messages first and pass the resulting UIDs to the bulk operation.",
      );
    }
    validateFolderPath(folder);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const query = this.buildSearchQuery(criteria);
        const result = await client.search(query, { uid: true });
        return result || [];
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Resolve `match` criteria to UIDs via SEARCH, then re-validate the survivors
   * with a FETCH-based existence pass. Filters out phantom UIDs that SEARCH
   * still surfaces after a recent move/delete — Proton Mail Bridge's SEARCH
   * index lags FETCH by ~1–2s after a mutation, so a back-to-back dry-run on
   * the same criteria in the same session can list UIDs that have already
   * left the folder.
   *
   * Used by bulk-op dry-run paths in `src/index.ts` so previews converge with
   * the FETCH-authoritative semantics of the live bulk operations. Without
   * this pipeline the preview overstated impact and listed cascade victims as
   * still-present, which is the primary reason an agent runs a dry-run in the
   * first place.
   */
  async resolveAndFilterUidsFromCriteria(folder: string, criteria: SearchCriteria): Promise<number[]> {
    const searchResults = await this.resolveUidsFromCriteria(folder, criteria);
    if (searchResults.length === 0) return [];
    const { existing } = await this.filterExistingUids(folder, searchResults);
    return existing;
  }

  /**
   * Walk a message body structure and collect attachment-like parts.
   *
   * Two modes:
   *  - `strict: false` (default) — treats both `Content-Disposition: attachment`
   *    and non-text `inline` parts (typically inline images) as attachments.
   *    Used by `list_attachments` / `download_attachment` / `read_message`
   *    so callers can still extract inline images embedded in HTML newsletters.
   *  - `strict: true` — only emits parts with `Content-Disposition: attachment`.
   *    Used by `searchMessages`'s `hasAttachment` post-filter so newsletters
   *    with inline banners aren't reported as carrying "attachments" in the
   *    user-intuitive sense (without it, inbox newsletters with banner images
   *    are falsely flagged as attachment-bearing).
   */
  /**
   * Estimate the decoded (real file) size from the IMAP-reported encoded octet
   * count. IMAP reports the size of the part *as encoded*. For base64 that's
   * inflated: standard MIME wraps base64 at 76 chars + CRLF (78 chars/line) and
   * every 4 encoded chars decode to 3 bytes, so decoded ≈ encoded × (76/78) ×
   * (3/4). **Approximate, not exact (±~2 bytes):** final-line padding and any
   * non-standard line wrapping shift the true count either direction — it is NOT
   * a strict floor (observed both +1 over and −2 under the saved size on real
   * Proton attachments). Close enough for display/budgeting, never byte-exact;
   * the real count only comes from actually decoding the part. Non-base64
   * encodings (7bit/8bit/binary, and ~quoted-printable) have encoded ≈ decoded,
   * so we pass them through unchanged.
   */
  static decodedSizeEstimate(encodedSize: number, encoding?: string): number {
    if (encoding?.toLowerCase() === "base64") {
      // Math.floor over the 4→3 ratio. This is an estimate, not exact: it lands
      // within ~2 bytes of the real decoded size but can fall on either side of
      // it (observed +1 over and −2 under on real attachments), since
      // the true count depends on padding and exact line wrapping we don't have
      // without the bytes. Close enough for display/budgeting; surfaced with `~`.
      return Math.floor((encodedSize * 76 * 3) / (78 * 4));
    }
    return encodedSize;
  }

  private static extractAttachments(
    structure: MessageStructureObject,
    options: { strict?: boolean } = {},
  ): AttachmentMeta[] {
    const attachments: AttachmentMeta[] = [];
    const strict = options.strict ?? false;

    const isAttached =
      structure.disposition === "attachment" ||
      (!strict && structure.disposition === "inline" && structure.type && !structure.type.startsWith("text/"));

    if (isAttached) {
      const encodedSize = structure.size || 0;
      attachments.push({
        partNumber: structure.part || "",
        filename: structure.dispositionParameters?.filename || structure.parameters?.name || "unnamed",
        contentType: structure.type || "application/octet-stream",
        size: ImapService.decodedSizeEstimate(encodedSize, structure.encoding),
        encodedSize,
      });
    }

    if (structure.childNodes) {
      for (const child of structure.childNodes) {
        attachments.push(...ImapService.extractAttachments(child, options));
      }
    }

    return attachments;
  }

  /**
   * Download an attachment by part number. Returns base64-encoded content.
   *
   * Pre-checks both the UID and the partNumber against the message's bodyStructure
   * so callers get actionable errors ("Part 42 not found on UID 7 ...; known parts: [2]")
   * instead of imapflow's opaque "Command failed" when the part doesn't exist.
   */
  async downloadAttachment(
    folder: string,
    uid: number,
    partNumber: string,
    options: { maxInlineBytes?: number } = {},
  ): Promise<{ content: string; contentType: string; filename: string }> {
    validateFolderPath(folder);
    validatePartNumber(partNumber);
    this.log(`[IMAP] Downloading attachment part ${partNumber} from UID ${uid} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        // Pre-check: fetch bodyStructure, verify UID exists and partNumber is a known attachment.
        const msg = await client.fetchOne(String(uid), { uid: true, bodyStructure: true }, { uid: true });
        if (!msg) {
          throw new Error(`Message UID ${uid} not found in ${folder}`);
        }
        const knownAttachments = msg.bodyStructure ? ImapService.extractAttachments(msg.bodyStructure) : [];
        const matchedPart = knownAttachments.find((a) => a.partNumber === partNumber);
        if (!matchedPart) {
          const known = knownAttachments.map((a) => a.partNumber).join(", ") || "(none)";
          throw new Error(`Part ${partNumber} not found on UID ${uid} in ${folder}; known parts: [${known}]`);
        }
        // Pre-flight size check: when the caller intends to return the bytes
        // inline (no saveTo), an oversized base64 payload will blow most agent
        // token caps. We compare the *encoded* size — that's what the inline
        // base64 response actually costs on the wire — not the decoded file
        // size. The threshold is approximate; the MCP transport applies its own
        // per-response token ceiling on top. Refuse before the download instead
        // of emitting base64 the caller's framework will reject.
        if (options.maxInlineBytes && matchedPart.encodedSize > options.maxInlineBytes) {
          throw new Error(
            `Attachment is ~${matchedPart.size} bytes (${matchedPart.encodedSize} bytes base64-encoded for inline transport, which exceeds the ${options.maxInlineBytes}-byte inline cap; the MCP framework's own per-response token cap may reject even smaller payloads). ` +
              `To download larger files, set ALLOW_FILE_DOWNLOAD_DIR in your MCP client config and pass \`saveTo\`. ` +
              `Example for Claude Desktop (claude_desktop_config.json): ` +
              `\`"env": { "ALLOW_FILE_DOWNLOAD_DIR": "/Users/you/proton-attachments" }\` — then call \`download_attachment\` with \`saveTo: "filename.ext"\` to write the bytes there and get the file path back instead of inline base64. ` +
              `Or call \`list_attachments\` first to confirm the size if you intentionally want the inline payload.`,
          );
        }

        const { meta, content } = await client.download(String(uid), partNumber, { uid: true });
        const chunks: Buffer[] = [];
        let totalSize = 0;
        for await (const chunk of content) {
          const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
          totalSize += buf.length;
          if (totalSize > this.maxAttachmentSize) {
            content.destroy?.();
            throw new Error(`Attachment exceeds maximum size of ${this.maxAttachmentSize} bytes`);
          }
          chunks.push(buf);
        }
        const fullBuffer = Buffer.concat(chunks);
        return {
          content: fullBuffer.toString("base64"),
          contentType: meta.contentType || "application/octet-stream",
          filename: meta.filename || "unnamed",
        };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Confirm a UID resolves to a message in the current mailbox lock.
   * IMAP STORE/DELETE/MOVE silently succeed against missing UIDs, so callers must
   * pre-check or risk reporting false success.
   */
  private async assertUidExists(client: ImapFlow, folder: string, uid: number): Promise<void> {
    const exists = await client.fetchOne(String(uid), { uid: true }, { uid: true });
    if (!exists) {
      throw new Error(`Message UID ${uid} not found in ${folder}`);
    }
  }

  /**
   * Resolve which of the requested UIDs currently exist in the open mailbox.
   *
   * Uses FETCH rather than SEARCH because IMAP SEARCH on Proton Mail Bridge can
   * lag the live mailbox state (server-side index propagation), which produced
   * spurious `notFound` entries even when the message was about to be successfully
   * mutated. FETCH operates on per-message state and only yields existing UIDs,
   * so the returned set is authoritative for accounting purposes.
   */
  private async existingUidSet(client: ImapFlow, uids: number[]): Promise<Set<number>> {
    if (uids.length === 0) return new Set();
    const existing = new Set<number>();
    for await (const msg of client.fetch(uids.join(","), { uid: true }, { uid: true })) {
      existing.add(msg.uid);
    }
    return existing;
  }

  /**
   * Partition an explicit UID list into `existing` and `notFound` using the
   * same FETCH-based pre-check that `bulkMove` / `bulkDelete` / `bulkUpdateFlags`
   * use at execution time. Used by the bulk-tool dry-run handlers so a preview
   * matches the live operation's `notFound` accounting — previously the dry-run
   * trusted the caller-supplied list verbatim, inflating the impact count and
   * omitting phantom UIDs.
   */
  async filterExistingUids(folder: string, uids: number[]): Promise<{ existing: number[]; notFound: number[] }> {
    validateFolderPath(folder);
    if (uids.length === 0) return { existing: [], notFound: [] };
    this.log(`[IMAP] Filtering ${uids.length} candidate UIDs in ${folder} for existence`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const existingSet = await this.existingUidSet(client, uids);
        return {
          existing: uids.filter((u) => existingSet.has(u)),
          notFound: uids.filter((u) => !existingSet.has(u)),
        };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Move a message to a different folder.
   */
  async moveMessage(folder: string, uid: number, destination: string): Promise<MoveResult> {
    validateFolderPath(folder);
    validateFolderPath(destination);
    this.log(`[IMAP] Moving UID ${uid} from ${folder} to ${destination}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        await this.assertUidExists(client, folder, uid);
        const result = await client.messageMove(String(uid), destination, { uid: true });
        if (result === false) {
          // imapflow's messageMove swallows the COPY error (copy.js:42-46) and just returns
          // false. To give the caller a meaningful message when the destination is missing,
          // probe with status() — it throws with mailboxMissing when the mailbox doesn't exist.
          try {
            await client.status(destination, { messages: true });
          } catch (err) {
            if (isMailboxMissingError(err)) {
              throw new Error(`Destination folder not found: ${destination}`);
            }
            // Probe itself failed for another reason — fall through to generic failure.
          }
          return { success: false, destination };
        }
        const newUid = result.uidMap?.get(uid);
        return { success: true, newUid, destination };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Permanently delete a message.
   */
  async deleteMessage(folder: string, uid: number): Promise<boolean> {
    validateFolderPath(folder);
    this.log(`[IMAP] Deleting UID ${uid} from ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        await this.assertUidExists(client, folder, uid);
        return await client.messageDelete(String(uid), { uid: true });
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Add and/or remove flags on a message, then verify the result with a post-STORE
   * FETCH and partition by what actually took effect.
   *
   * Proton Mail Bridge has been observed to silently drop user keywords (e.g.
   * "Important_Tag") on STORE: `messageFlagsAdd` returns success but the
   * keyword does not show up on subsequent FETCH. The post-STORE verify here
   * catches that and reports the unrequested-but-still-missing flags in
   * `notApplied`, mirroring the `update_message_labels` response shape.
   *
   * System flag \Recent is server-managed and will also be reported in
   * notApplied if a caller asks for it — that's the correct signal.
   */
  async updateFlags(
    folder: string,
    uid: number,
    flagsToAdd: string[],
    flagsToRemove: string[],
  ): Promise<{ added: string[]; removed: string[]; notApplied: string[] }> {
    validateFolderPath(folder);
    this.log(`[IMAP] Updating flags for UID ${uid} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        await this.assertUidExists(client, folder, uid);
        if (flagsToAdd.length > 0) {
          await client.messageFlagsAdd(String(uid), flagsToAdd, { uid: true });
        }
        if (flagsToRemove.length > 0) {
          await client.messageFlagsRemove(String(uid), flagsToRemove, { uid: true });
        }

        // Post-STORE verify: refetch flags and partition requested operations.
        const verified = await client.fetchOne(String(uid), { uid: true, flags: true }, { uid: true });
        const actual = new Set<string>(verified && verified.flags ? [...verified.flags] : []);
        const added: string[] = [];
        const removed: string[] = [];
        const notApplied: string[] = [];
        for (const f of flagsToAdd) (actual.has(f) ? added : notApplied).push(f);
        for (const f of flagsToRemove) (actual.has(f) ? notApplied : removed).push(f);
        return { added, removed, notApplied };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Add and/or remove Proton labels on a message.
   *
   * Proton labels live in the "Labels/" IMAP namespace and are additive (the
   * underlying message stays in its source folder). Mechanics:
   *  - Add a label: pre-check the label mailbox exists via `status()`, then
   *    IMAP `COPY` from the source folder to `Labels/<name>`. The source UID
   *    is unchanged; a new UID is assigned in the label mailbox.
   *  - Remove a label: locate the message in the label mailbox by its
   *    `Message-ID` header, then `messageDelete` that UID (removes only the
   *    label, leaves the message in its source folder intact).
   *
   * Adds are strict: a missing label throws `Label not found: <path>`. The
   * status() pre-check is required because imapflow's `messageCopy` returns
   * `false` (rather than throws) on a missing destination, and Proton Mail
   * Bridge returns a bare "Command failed" without the `mailboxMissing`
   * signal — both of which would silently swallow the add otherwise.
   *
   * Removals are idempotent: removing a label that isn't applied, or that
   * doesn't exist as a mailbox at all, is a no-op. The return value reports
   * which removes did vs. didn't take effect so callers don't mislead users.
   */
  async updateLabels(
    folder: string,
    uid: number,
    labelsToAdd: string[],
    labelsToRemove: string[],
  ): Promise<{ added: string[]; removed: string[]; notApplied: string[] }> {
    validateFolderPath(folder);
    for (const l of [...labelsToAdd, ...labelsToRemove]) {
      if (!/^Labels\//.test(l)) {
        throw new Error(`Label paths must start with "Labels/" (got: "${l}")`);
      }
      validateFolderPath(l);
    }
    this.log(`[IMAP] Updating labels for UID ${uid} in ${folder} (+${labelsToAdd.length}, -${labelsToRemove.length})`);

    const client = this.createClient();
    try {
      await client.connect();

      // Resolve Message-ID up front — removals need it to find the message in
      // each label mailbox (per-mailbox UIDs differ from the source). UID
      // check before label pre-check so a wrong UID surfaces ahead of label
      // errors (matches the update_message_flags / move_message pattern).
      let messageId = "";
      const sourceLock = await this.lockFolder(client, folder);
      try {
        const msg = await client.fetchOne(String(uid), { uid: true, envelope: true }, { uid: true });
        if (!msg) throw new Error(`Message UID ${uid} not found in ${folder}`);
        messageId = msg.envelope?.messageId ?? "";

        // Pre-check that each label-to-add EXISTS before we COPY. status()
        // throws with mailboxMissing when the label isn't a mailbox; we
        // translate to "Label not found". Required because imapflow's
        // messageCopy returns `false` (rather than throws) for missing
        // destinations, and Proton Mail Bridge returns a bare "Command
        // failed" without `mailboxMissing` — both would silently swallow.
        for (const label of labelsToAdd) {
          try {
            await client.status(label, { messages: true });
          } catch (err) {
            if (isMailboxMissingError(err)) {
              throw new Error(`Label not found: ${label}`);
            }
            throw err;
          }
        }

        for (const label of labelsToAdd) {
          const result = await client.messageCopy(String(uid), label, { uid: true });
          if (result === false) {
            // Pre-check confirmed the mailbox existed; a falsy result here is
            // a genuine COPY failure (e.g. permissions, quota), not "missing".
            throw new Error(`Failed to copy UID ${uid} to ${label}`);
          }
        }
      } finally {
        sourceLock.release();
      }

      if (labelsToRemove.length > 0 && !messageId) {
        throw new Error(`Cannot remove labels — message UID ${uid} in ${folder} has no Message-ID header`);
      }

      const removed: string[] = [];
      const notApplied: string[] = [];
      for (const label of labelsToRemove) {
        let labelLock;
        try {
          labelLock = await this.lockFolder(client, label);
        } catch (err) {
          // Label doesn't exist as a mailbox → the message can't carry it → no-op.
          if (err instanceof Error && /^Folder not found:/.test(err.message)) {
            notApplied.push(label);
            continue;
          }
          throw err;
        }
        try {
          const matches = await client.search({ header: { "Message-ID": messageId } }, { uid: true });
          if (!matches || matches.length === 0) {
            notApplied.push(label);
            continue;
          }
          await client.messageDelete(matches.join(","), { uid: true });
          removed.push(label);
        } finally {
          labelLock.release();
        }
      }

      return { added: [...labelsToAdd], removed, notApplied };
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Return attachment metadata for a message without downloading the body.
   */
  async listAttachments(folder: string, uid: number): Promise<AttachmentMeta[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Listing attachments for UID ${uid} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const msg = await client.fetchOne(String(uid), { uid: true, bodyStructure: true }, { uid: true });
        if (!msg) {
          throw new Error(`Message UID ${uid} not found in ${folder}`);
        }
        return msg.bodyStructure ? ImapService.extractAttachments(msg.bodyStructure) : [];
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Mark all unread messages in a folder as read.
   * Optionally limit to messages older than a given date.
   * Returns the number of messages marked.
   */
  async markAllRead(folder: string, olderThan?: string, options: { dryRun?: boolean } = {}): Promise<number> {
    validateFolderPath(folder);
    const previewSuffix = options.dryRun ? " (dry-run)" : "";
    this.log(
      `[IMAP] Marking all unread as read in ${folder}${olderThan ? ` before ${olderThan}` : ""}${previewSuffix}`,
    );
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const query: SearchObject = { seen: false };
        if (olderThan) query.before = olderThan;

        const result = await client.search(query, { uid: true });
        if (!result || result.length === 0) return 0;

        if (options.dryRun) return result.length;

        const uidRange = result.join(",");
        await client.messageFlagsAdd(uidRange, ["\\Seen"], { uid: true });
        return result.length;
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Find a message by its Message-ID header. Returns the UID if found.
   */
  async findByMessageId(folder: string, messageId: string): Promise<number | undefined> {
    validateFolderPath(folder);
    this.log(`[IMAP] Searching for Message-ID ${messageId} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const result = await client.search({ header: { "Message-ID": messageId } }, { uid: true });
        if (!result || result.length === 0) return undefined;
        return result[result.length - 1]; // Return the last (most recent) match
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Look up the Sent-folder copy of a just-sent message with a short retry loop
   * (Proton's SEARCH index can lag a fresh APPEND by a few seconds), and, when
   * found, also FETCH its Reply-To header so callers can detect server-side
   * rewriting (Proton SMTP overrides Reply-To values not matching an
   * authenticated identity).
   *
   * The Sent folder path is resolved via the IMAP `\Sent` special-use annotation,
   * falling back to the literal "Sent" name if no annotation is published.
   *
   * Returns `uid: undefined` when the lookup didn't converge — the caller should
   * surface that as a best-effort miss, not an error.
   */
  async findSentCopyMeta(
    messageId: string,
    options: { retries?: number; backoffMs?: number[] } = {},
  ): Promise<{ uid?: number; folder: string; replyTo?: string }> {
    // Default to 8 attempts with 7 waits totaling ~30s — Proton Bridge SEARCH
    // lag occasionally exceeded the prior ~17s window under burst conditions,
    // producing non-deterministic verification (back-to-back send_email calls
    // in the same session saw one verified and the next unverified). The wait
    // happens between attempts, so attempts=8 uses backoff[0..6] (the trailing
    // value would fire only if a 9th attempt existed). The loop breaks as soon
    // as the Message-ID resolves, so a fast-converging Sent copy still returns
    // in ~1s — the long tail only kicks in when the bridge actually lags.
    const retries = options.retries ?? 8;
    const backoff = options.backoffMs ?? [500, 1000, 2000, 3500, 5500, 7500, 10000];
    this.log(`[IMAP] Looking up Sent copy for Message-ID ${messageId} (retries=${retries})`);

    const client = this.createClient();
    try {
      await client.connect();
      const folders = await this.specialFolderResolver.resolve(client);
      const folder = folders.sent ?? "Sent";

      let uid: number | undefined;
      for (let attempt = 0; attempt < retries; attempt++) {
        try {
          const lock = await this.lockFolder(client, folder);
          try {
            const result = await client.search({ header: { "Message-ID": messageId } }, { uid: true });
            if (result && result.length > 0) {
              uid = result[result.length - 1];
              break;
            }
          } finally {
            lock.release();
          }
        } catch (err) {
          if (!isMailboxMissingError(err)) throw err;
        }
        if (attempt < retries - 1) {
          await new Promise((r) => setTimeout(r, backoff[Math.min(attempt, backoff.length - 1)]));
        }
      }

      if (uid === undefined) return { folder };

      // Found — also read the delivered Reply-To header so callers can detect rewriting.
      let replyTo: string | undefined;
      try {
        const lock = await this.lockFolder(client, folder);
        try {
          const msg = await client.fetchOne(String(uid), { uid: true, headers: ["reply-to"] }, { uid: true });
          const headersStr = msg && msg.headers ? msg.headers.toString("utf-8") : "";
          const match = headersStr.match(/^Reply-To:\s*(.+(?:\r?\n[ \t]+.+)*)/im);
          if (match) replyTo = match[1].replace(/\r?\n[ \t]+/g, " ").trim();
        } finally {
          lock.release();
        }
      } catch {
        // Reply-To inspection is best-effort; don't fail the whole lookup.
      }

      return { uid, folder, replyTo };
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Get all messages in a conversation thread by walking References and In-Reply-To headers.
   * Searches within a single folder. Use `getThreadByMessageId` for cross-folder threads.
   */
  async getThread(folder: string, uid: number, limit: number): Promise<MessageSummary[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Getting thread for UID ${uid} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const seed = await this.fetchThreadSeed(client, uid, folder);
        const knownIds = this.collectSeedIds(seed);
        const allUids = new Set<number>([uid]);
        await this.walkReferences(client, knownIds, allUids, limit);
        return await this.fetchThreadEnvelopes(client, allUids, limit);
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Walk a conversation thread by Message-ID across the default folder set
   * (INBOX + Sent + All Mail). Avoids the UID-collision footgun of `getThread`:
   * Message-IDs are globally unique, UIDs are per-folder.
   *
   * Missing folders (e.g. "All Mail" on providers that don't ship it) are skipped.
   * Returns messages tagged with their source folder so the caller can navigate back.
   */
  async getThreadByMessageId(
    messageId: string,
    limit: number,
    folders: readonly string[] = DEFAULT_THREAD_FOLDERS,
  ): Promise<MessageSummary[]> {
    for (const folder of folders) validateFolderPath(folder);
    this.log(`[IMAP] Getting thread by Message-ID ${messageId} across ${folders.join(", ")}`);
    const client = this.createClient();
    let threadRows: (MessageSummary & { folder: string; mailboxCopies?: number; otherFolders?: string[] })[] = [];
    try {
      await client.connect();

      // Step 1: locate the seed message by Message-ID. Walk folders until found.
      let seedFolder: string | undefined;
      let seedUid: number | undefined;
      for (const folder of folders) {
        let lock;
        try {
          lock = await client.getMailboxLock(folder);
        } catch (err) {
          if (isMailboxMissingError(err)) continue;
          throw err;
        }
        try {
          const found = await client.search({ header: { "Message-ID": messageId } }, { uid: true });
          if (found && found.length > 0) {
            seedFolder = folder;
            seedUid = found[found.length - 1];
            break;
          }
        } finally {
          lock.release();
        }
      }

      if (seedFolder === undefined || seedUid === undefined) {
        throw new Error(`Message with Message-ID ${messageId} not found in ${folders.join(", ")}`);
      }

      // Step 2: fetch the seed headers to find related Message-IDs.
      const seedLock = await this.lockFolder(client, seedFolder);
      let knownIds: Set<string>;
      try {
        const seed = await this.fetchThreadSeed(client, seedUid, seedFolder);
        knownIds = this.collectSeedIds(seed);
      } finally {
        seedLock.release();
      }

      // Step 3: walk each folder collecting UIDs that match any known Message-ID.
      const perFolder = new Map<string, Set<number>>();
      for (const folder of folders) {
        let lock;
        try {
          lock = await client.getMailboxLock(folder);
        } catch (err) {
          if (isMailboxMissingError(err)) continue;
          throw err;
        }
        try {
          const uids = new Set<number>();
          if (folder === seedFolder) uids.add(seedUid);
          await this.walkReferences(client, knownIds, uids, limit);
          if (uids.size > 0) perFolder.set(folder, uids);
        } finally {
          lock.release();
        }
      }

      // Step 4: fetch envelopes per-folder and merge.
      const merged: (MessageSummary & { folder: string })[] = [];
      for (const [folder, uids] of perFolder.entries()) {
        const lock = await this.lockFolder(client, folder);
        try {
          const envelopes = await this.fetchThreadEnvelopes(client, uids, limit);
          for (const env of envelopes) merged.push({ ...env, folder });
        } finally {
          lock.release();
        }
      }

      // De-duplicate by Message-ID: under Proton's label model, one physical
      // message can appear in multiple mailboxes (INBOX + All Mail, etc.), each
      // with its own UID. The headline thread count should reflect distinct
      // messages, not mailbox copies. Prefer the non-"All Mail" copy as the
      // canonical row; surface the alternate folders via mailboxCopies/otherFolders
      // so the caller still sees the cross-namespace state.
      const byMsgId = new Map<string, (MessageSummary & { folder: string; otherFolders?: string[] })[]>();
      const noMsgId: (MessageSummary & { folder: string })[] = [];
      for (const row of merged) {
        const id = row.messageId;
        if (!id) {
          noMsgId.push(row);
          continue;
        }
        const list = byMsgId.get(id) ?? [];
        list.push(row);
        byMsgId.set(id, list);
      }
      const deduped: (MessageSummary & { folder: string; mailboxCopies?: number; otherFolders?: string[] })[] = [];
      for (const rows of byMsgId.values()) {
        // Stable preference: non-"All Mail" first, then by UID for determinism.
        rows.sort((a, b) => {
          const aAll = /^all mail$/i.test(a.folder) ? 1 : 0;
          const bAll = /^all mail$/i.test(b.folder) ? 1 : 0;
          if (aAll !== bAll) return aAll - bAll;
          return a.uid - b.uid;
        });
        const [primary, ...rest] = rows;
        deduped.push({
          ...primary,
          mailboxCopies: rows.length,
          otherFolders: rest.map((r) => r.folder),
        });
      }
      const result = [...deduped, ...noMsgId];
      result.sort((a, b) => {
        const da = a.date ? new Date(a.date).getTime() : 0;
        const db = b.date ? new Date(b.date).getTime() : 0;
        return da !== db ? da - db : a.uid - b.uid;
      });
      threadRows = result.slice(0, limit);
    } finally {
      await client.logout().catch(() => {});
    }
    // Rewrite any All Mail-only members to their real storage folder + UID. The
    // default walk (INBOX + Sent + All Mail) misses threads stored in a user
    // folder (e.g. Folders/X) — those surface only via the All Mail virtual copy,
    // pairing the row with an All Mail UID that is invalid in any real folder. An
    // agent copying that UID into a single-folder tool (move_message /
    // delete_message) would then target the wrong message or no-op. Reuses the
    // same reroute the thread-mutation ops use; it opens a short-lived connection
    // only when orphans are actually present, and is a no-op otherwise.
    return await this.rerouteAllMailOrphans(threadRows);
  }

  private async fetchThreadSeed(client: ImapFlow, uid: number, folder: string): Promise<FetchMessageObject> {
    const seed = await client.fetchOne(
      String(uid),
      { uid: true, envelope: true, headers: ["references"] },
      { uid: true },
    );
    if (!seed) throw new Error(`Message UID ${uid} not found in ${folder}`);
    if (!seed.envelope?.messageId) throw new Error(`Message UID ${uid} has no Message-ID`);
    return seed;
  }

  private collectSeedIds(seed: FetchMessageObject): Set<string> {
    const ids = new Set<string>();
    if (seed.envelope?.messageId) ids.add(seed.envelope.messageId);
    const headersStr = seed.headers ? seed.headers.toString("utf-8") : "";
    const referencesMatch = headersStr.match(/^References:\s*(.+(?:\r?\n[ \t]+.+)*)/im);
    const referencesRaw = referencesMatch ? referencesMatch[1].replace(/\r?\n[ \t]+/g, " ") : "";
    const refMatches = referencesRaw.match(/<[^>]+>/g);
    if (refMatches) for (const ref of refMatches) ids.add(ref);
    if (seed.envelope?.inReplyTo) ids.add(seed.envelope.inReplyTo);
    return ids;
  }

  private async walkReferences(
    client: ImapFlow,
    knownIds: Set<string>,
    collected: Set<number>,
    limit: number,
  ): Promise<void> {
    for (const msgId of knownIds) {
      const byId = await client.search({ header: { "Message-ID": msgId } }, { uid: true });
      if (byId && byId.length > 0) for (const u of byId) collected.add(u);
      const byRef = await client.search({ header: { References: msgId } }, { uid: true });
      if (byRef && byRef.length > 0) for (const u of byRef) collected.add(u);
      const byReply = await client.search({ header: { "In-Reply-To": msgId } }, { uid: true });
      if (byReply && byReply.length > 0) for (const u of byReply) collected.add(u);
      if (collected.size >= limit) break;
    }
  }

  private async fetchThreadEnvelopes(client: ImapFlow, uids: Set<number>, limit: number): Promise<MessageSummary[]> {
    if (uids.size === 0) return [];
    const uidList = [...uids].slice(0, limit);
    const fetchRange = uidList.join(",");
    const messages: MessageSummary[] = [];
    for await (const msg of client.fetch(fetchRange, { uid: true, envelope: true, flags: true }, { uid: true })) {
      messages.push(this.toSummary(msg));
    }
    messages.sort((a, b) => {
      const da = a.date ? new Date(a.date).getTime() : 0;
      const db = b.date ? new Date(b.date).getTime() : 0;
      return da !== db ? da - db : a.uid - b.uid;
    });
    return messages;
  }

  /**
   * Move multiple messages to a destination folder in one IMAP operation.
   * Pre-checks existence via FETCH (see existingUidSet) to identify phantom UIDs
   * before issuing messageMove.
   */
  async bulkMove(folder: string, uids: number[], destination: string): Promise<BulkMoveResult> {
    validateFolderPath(folder);
    validateFolderPath(destination);
    this.log(`[IMAP] Bulk-moving ${uids.length} UIDs from ${folder} to ${destination}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const existingSet = await this.existingUidSet(client, uids);
        const notFound = uids.filter((u) => !existingSet.has(u));
        const toMove = uids.filter((u) => existingSet.has(u));

        if (toMove.length === 0) {
          return { moved: 0, notFound, newUids: [], destination };
        }

        const result = await client.messageMove(toMove.join(","), destination, { uid: true });
        if (result === false) {
          try {
            await client.status(destination, { messages: true });
          } catch (err) {
            if (isMailboxMissingError(err)) {
              throw new Error(`Destination folder not found: ${destination}`);
            }
          }
          return { moved: 0, notFound: uids, newUids: [], destination };
        }

        const newUids = toMove.map((u) => result.uidMap?.get(u)).filter((v): v is number => typeof v === "number");
        return { moved: toMove.length, notFound, newUids, destination };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Delete multiple messages in one IMAP operation.
   * When permanent=false, moves to the resolved Trash folder (soft-delete).
   * When permanent=true, expunges via messageDelete.
   */
  async bulkDelete(folder: string, uids: number[], permanent: boolean): Promise<BulkDeleteResult> {
    validateFolderPath(folder);
    this.log(`[IMAP] Bulk-deleting ${uids.length} UIDs from ${folder} (permanent=${permanent})`);

    if (!permanent) {
      const special = await this.getSpecialFolders();
      const trashFolder = special.trash ?? "Trash";
      const move = await this.bulkMove(folder, uids, trashFolder);
      return {
        deleted: move.moved,
        notFound: move.notFound,
        newUids: move.newUids,
        destination: move.destination,
      };
    }

    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const existingSet = await this.existingUidSet(client, uids);
        const notFound = uids.filter((u) => !existingSet.has(u));
        const toDelete = uids.filter((u) => existingSet.has(u));

        if (toDelete.length === 0) {
          return { deleted: 0, notFound };
        }

        await client.messageDelete(toDelete.join(","), { uid: true });
        return { deleted: toDelete.length, notFound };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Add and/or remove flags on multiple messages in one IMAP operation.
   * Pre-checks existence via FETCH (see existingUidSet) to identify phantom UIDs
   * without being misled by SEARCH index lag.
   */
  async bulkUpdateFlags(
    folder: string,
    uids: number[],
    flagsToAdd: string[],
    flagsToRemove: string[],
  ): Promise<BulkFlagResult> {
    validateFolderPath(folder);
    this.log(`[IMAP] Bulk flag update on ${uids.length} UIDs in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const existingSet = await this.existingUidSet(client, uids);
        const notFound = uids.filter((u) => !existingSet.has(u));
        const toUpdate = uids.filter((u) => existingSet.has(u));

        if (toUpdate.length === 0) {
          return { affected: 0, notFound, notApplied: [] };
        }

        const range = toUpdate.join(",");
        if (flagsToAdd.length > 0) {
          await client.messageFlagsAdd(range, flagsToAdd, { uid: true });
        }
        if (flagsToRemove.length > 0) {
          await client.messageFlagsRemove(range, flagsToRemove, { uid: true });
        }

        // Post-STORE verify: re-fetch flags and compute the silently-dropped set.
        // A flag is reported notApplied if the server failed to apply (or
        // remove) it on *every* affected UID — the common Proton case is a
        // user keyword that's universally dropped. Per-UID partial outcomes
        // would surface as `affected` < `toUpdate.length`, which v1.0 does not
        // try to model.
        const flagsByUid = new Map<number, Set<string>>();
        for await (const m of client.fetch(range, { uid: true, flags: true }, { uid: true })) {
          flagsByUid.set(m.uid, new Set(m.flags ? [...m.flags] : []));
        }
        const notApplied: string[] = [];
        for (const f of flagsToAdd) {
          if ([...flagsByUid.values()].every((s) => !s.has(f))) notApplied.push(f);
        }
        for (const f of flagsToRemove) {
          if ([...flagsByUid.values()].every((s) => s.has(f))) notApplied.push(f);
        }
        return { affected: toUpdate.length, notFound, notApplied };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Add and/or remove Proton labels on many messages in one connection.
   *
   * Adds are batched: one IMAP COPY of the joined UID range per label. Removes
   * are per-Message-ID lookups inside the label mailbox — labels carry their
   * own per-mailbox UIDs that differ from the source UIDs, so we must locate
   * each via the source's Message-ID header.
   *
   * `notApplied` aggregates labels whose operation was a no-op across the whole
   * batch (e.g. a remove for a label none of the affected UIDs actually
   * carried, or a remove targeting a label that doesn't exist as a mailbox).
   */
  async bulkUpdateLabels(
    folder: string,
    uids: number[],
    labelsToAdd: string[],
    labelsToRemove: string[],
  ): Promise<BulkLabelResult> {
    validateFolderPath(folder);
    for (const l of [...labelsToAdd, ...labelsToRemove]) {
      if (!/^Labels\//.test(l)) {
        throw new Error(`Label paths must start with "Labels/" (got: "${l}")`);
      }
      validateFolderPath(l);
    }
    this.log(`[IMAP] Bulk label update on ${uids.length} UIDs in ${folder}`);

    const client = this.createClient();
    try {
      await client.connect();

      // Lock the source folder BEFORE the existence pre-check — `existingUidSet`
      // runs an IMAP FETCH, which requires a selected mailbox. Without the
      // lock, FETCH operates on whatever mailbox imapflow happened to have
      // selected (or none), so every UID would be reported as `notFound` while
      // the subsequent COPY still succeeded — making bulk_update_labels report
      // "0 message(s) updated" even though the label was applied.
      const messageIds = new Map<number, string>();
      let notFound: number[] = [];
      let toUpdate: number[] = [];
      const sourceLock = await this.lockFolder(client, folder);
      try {
        const existingSet = await this.existingUidSet(client, uids);
        notFound = uids.filter((u) => !existingSet.has(u));
        toUpdate = uids.filter((u) => existingSet.has(u));

        if (toUpdate.length === 0) {
          return { affected: 0, notFound, notApplied: [] };
        }

        // Fetch Message-IDs for the affected UIDs in the source folder — required
        // for the remove path.
        for await (const m of client.fetch(toUpdate.join(","), { uid: true, envelope: true }, { uid: true })) {
          if (m.envelope?.messageId) messageIds.set(m.uid, m.envelope.messageId);
        }

        // ADD path: status() pre-check per label, then one COPY of the range.
        for (const label of labelsToAdd) {
          try {
            await client.status(label, { messages: true });
          } catch (err) {
            if (isMailboxMissingError(err)) throw new Error(`Label not found: ${label}`);
            throw err;
          }
        }
        for (const label of labelsToAdd) {
          const result = await client.messageCopy(toUpdate.join(","), label, { uid: true });
          if (result === false) {
            throw new Error(`Failed to copy ${toUpdate.length} UID(s) to ${label}`);
          }
        }
      } finally {
        sourceLock.release();
      }

      // REMOVE path: per-label, walk the label mailbox and DELETE matches.
      const notApplied: string[] = [];
      for (const label of labelsToRemove) {
        let labelLock;
        try {
          labelLock = await this.lockFolder(client, label);
        } catch (err) {
          if (err instanceof Error && /^Folder not found:/.test(err.message)) {
            notApplied.push(label);
            continue;
          }
          throw err;
        }
        try {
          const toDelete: number[] = [];
          for (const mid of messageIds.values()) {
            const matches = await client.search({ header: { "Message-ID": mid } }, { uid: true });
            if (matches && matches.length > 0) toDelete.push(...matches);
          }
          if (toDelete.length === 0) {
            notApplied.push(label);
            continue;
          }
          await client.messageDelete(toDelete.join(","), { uid: true });
        } finally {
          labelLock.release();
        }
      }

      return { affected: toUpdate.length, notFound, notApplied };
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Save a raw RFC 5322 message to a folder (IMAP APPEND).
   * Returns the UID of the appended message if the server supports UIDPLUS.
   */
  async saveDraft(folder: string, rawMessage: Buffer): Promise<{ uid?: number }> {
    validateFolderPath(folder);
    this.log(`[IMAP] Saving draft to ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      let result;
      try {
        result = await client.append(folder, rawMessage, ["\\Draft", "\\Seen"]);
      } catch (err) {
        if (isMailboxMissingError(err)) {
          throw new Error(`Folder not found: ${folder}`);
        }
        throw err;
      }
      if (result === false) {
        throw new Error(`Failed to save draft to ${folder}`);
      }
      return { uid: result.uid };
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Raw mailbox-create call shared by `createFolder` and `createLabel`. Has no
   * namespace policy — both wrappers apply their own front-side guards. Translates
   * already-exists into idempotent success (covers both the `ALREADYEXISTS` code
   * path and Proton Mail Bridge's bare "Command failed" via a status() probe).
   *
   * `friendlyKind` controls the surfaced wording in the namespace-hint fallback
   * ("folder" vs "label") for unprefixed paths that Proton refuses outright.
   */
  private async createMailboxRaw(
    path: string,
    friendlyKind: "folder" | "label",
  ): Promise<{ created: boolean; alreadyExists?: boolean }> {
    this.log(`[IMAP] Creating ${friendlyKind} ${path}`);
    const client = this.createClient();
    try {
      await client.connect();
      try {
        await client.mailboxCreate(path);
        return { created: true };
      } catch (err) {
        const e = err as Error & { serverResponseCode?: string };
        if (e.serverResponseCode === "ALREADYEXISTS" || /already exists/i.test(e.message)) {
          return { created: false, alreadyExists: true };
        }
        try {
          await client.status(path, { messages: true });
          return { created: false, alreadyExists: true };
        } catch {
          // Probe failed — the mailbox really doesn't exist; the create
          // genuinely failed (not an idempotent already-exists).
        }
        // The create genuinely failed with a bare "Command failed" (no
        // actionable code). Consult the live mailbox list to disambiguate the
        // real cause at runtime rather than asserting an unverifiable Proton
        // rule. One extra round-trip on the error path is cheap.
        let existing: Array<{ path?: string }> = [];
        try {
          const listed = await client.list();
          if (Array.isArray(listed)) existing = listed;
        } catch {
          // list() unavailable — fall through to the generic message.
        }
        const leaf = (path.split("/").pop() ?? path).toLowerCase();

        // (1) The exact path already exists but the status() probe didn't see it
        //     (STATUS can behave differently on Proton label-backed mailboxes) —
        //     treat as idempotent success, matching the create* contract.
        if (existing.some((mb) => (mb.path ?? "").toLowerCase() === path.toLowerCase())) {
          return { created: false, alreadyExists: true };
        }

        if (!/^(Folders|Labels)\//i.test(path)) {
          const hint = friendlyKind === "label" ? "Labels/" : "Folders/";
          throw new Error(
            `Failed to create ${friendlyKind} "${path}". On Proton Mail, ${friendlyKind}s must be created under the "${hint}" namespace (e.g. "${hint}${path}"). Underlying error: ${e.message}`,
          );
        }

        // (2) A different mailbox already owns this leaf name. Proton enforces
        //     unique names across folders and labels, so a label/folder can't
        //     reuse a name held by another mailbox (e.g. the system "Archive"
        //     folder) — and reports the conflict only as a bare "Command
        //     failed". This is evidence-based: we only claim a collision when
        //     the list actually shows one, and it correctly covers the case an
        //     agent hit (Labels/Important/Test succeed, Labels/Archive fails).
        const collision = existing.find((mb) => (mb.path?.split("/").pop() ?? "").toLowerCase() === leaf);
        if (collision) {
          throw new Error(
            `Failed to create ${friendlyKind} "${path}": a mailbox named "${path.split("/").pop()}" already exists ("${collision.path}"), and Proton won't let a ${friendlyKind} reuse a name held by another mailbox. Choose a different name.`,
          );
        }

        // (3) No collision found — translate the opaque bridge error into a
        //     generic-but-actionable message instead of the raw passthrough.
        throw new Error(
          `Failed to create ${friendlyKind} "${path}": Proton rejected the request (server said: "${e.message}"). Likely causes: the name collides with a reserved/existing mailbox, contains characters Proton disallows, or a parent segment doesn't exist.`,
        );
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  async createFolder(path: string): Promise<{ created: boolean; alreadyExists?: boolean }> {
    validateFolderPath(path);
    // Reject `.` and `..` segments at creation time. Previously create_folder
    // accepted them (Proton happily makes a literal `Folders/..` mailbox) but
    // delete_folder refused to clean them up, leaving the user stuck with
    // folders only removable via the Proton web UI. Now both ends apply the
    // same segment guard.
    if (path.split("/").some((seg) => seg === "." || seg === "..")) {
      throw new Error(`Folder path contains invalid segments (got: "${path}").`);
    }
    // create_folder is for the Folders/ namespace. Labels live under Labels/
    // but the create surface for them is the dedicated `create_label` tool,
    // which takes a bare label name and prepends the namespace internally.
    // Routing folder-creation through Labels/ silently muddies tool discovery
    // for agents — refuse and point at the right tool.
    if (/^Labels\//i.test(path)) {
      throw new Error(
        `Use the create_label tool to create Proton labels — it takes a bare name like "Important" and prepends Labels/ internally (got: "${path}").`,
      );
    }
    return this.createMailboxRaw(path, "folder");
  }

  /**
   * Create a Proton label from a bare name (e.g. "Important"). Prepends
   * "Labels/" internally and bypasses createFolder's Labels/-rejection guard
   * (which is intended to redirect callers of `create_folder` to *this* tool —
   * we'd circular-loop if we routed through createFolder).
   */
  async createLabel(name: string): Promise<{ created: boolean; alreadyExists?: boolean; path: string }> {
    if (!name || name.includes("/")) {
      throw new Error(`Label name must be a non-empty bare name without "/" (got: "${name}").`);
    }
    if (name.split("/").some((seg) => seg === "." || seg === "..")) {
      throw new Error(`Label name contains invalid segments (got: "${name}").`);
    }
    const path = `Labels/${name}`;
    validateFolderPath(path);
    const result = await this.createMailboxRaw(path, "label");
    return { ...result, path };
  }

  async renameFolder(from: string, to: string): Promise<void> {
    validateFolderPath(from);
    validateFolderPath(to);
    // Mirror deleteFolder's namespace + segment guards on BOTH ends. Without
    // this, `rename_folder(from: "INBOX", to: "Folders/Hacked")` succeeded —
    // Proton accepts the rename and relocates the contents, leaving INBOX
    // empty (it gets auto-recreated by Proton, but the original UIDs are gone
    // and the rename is not in-tool reversible). delete_folder has had this
    // check since v0.6.0; rename_folder shipped without it. Required because
    // imapflow's mailboxRename doesn't refuse system mailboxes by name, and
    // Proton Mail only refuses some (Sent, Trash) but not INBOX.
    for (const [label, path] of [
      ["from", from],
      ["to", to],
    ] as const) {
      if (!/^(Folders|Labels)\//i.test(path)) {
        throw new Error(
          `rename_folder is restricted to the "Folders/" and "Labels/" namespaces to protect system mailboxes (${label}: "${path}").`,
        );
      }
      if (path.split("/").some((seg) => seg === "." || seg === "..")) {
        throw new Error(`Folder path contains invalid segments (${label}: "${path}").`);
      }
    }
    // Cross-namespace renames (Folders/X → Labels/Y or vice versa) are
    // semantically nonsense on Proton — labels and folders are different
    // primitives (exclusive vs additive). imapflow forwards the request and
    // Proton responds with a bare "Command failed" that leaks through. Catch
    // the mismatch upfront with an actionable error so the caller knows to
    // copy messages explicitly instead.
    const fromNs = from.match(/^(Folders|Labels)\//i)?.[1].toLowerCase();
    const toNs = to.match(/^(Folders|Labels)\//i)?.[1].toLowerCase();
    if (fromNs && toNs && fromNs !== toNs) {
      throw new Error(
        `rename_folder requires the same namespace on both ends; converting between Folder and Label is not supported (got from: "${from}", to: "${to}"). Use bulk_move (for messages) or update_message_labels (to retag) explicitly.`,
      );
    }
    this.log(`[IMAP] Renaming folder ${from} → ${to}`);
    const client = this.createClient();
    try {
      await client.connect();
      try {
        await client.mailboxRename(from, to);
      } catch (err) {
        if (isMailboxMissingError(err)) {
          throw new Error(`Folder not found: ${from}`);
        }
        // Proton Mail Bridge returns a bare "Command failed" without the
        // mailboxMissing signal when the source doesn't exist. Probe with
        // status() — mirrors the post-failure probe pattern moveMessage uses.
        // Post-failure (not pre-check) avoids an extra round-trip on success
        // and the TOCTOU window where the source vanishes between checks.
        try {
          await client.status(from, { messages: true });
        } catch (probeErr) {
          if (isMailboxMissingError(probeErr)) {
            throw new Error(`Folder not found: ${from}`);
          }
        }
        throw err;
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Delete a folder or label container. Restricted to the "Folders/" and
   * "Labels/" namespaces so system mailboxes (INBOX, Sent, Trash, etc.) cannot
   * be removed by accident.
   *
   * On Proton Mail Bridge this is *not* a destructive message operation:
   * deleting a folder relocates its contents into "All Mail"; deleting a label
   * simply removes the label and leaves the underlying message untouched in
   * its source folder. The risk profile is metadata-only, so we don't require
   * a confirmation flag.
   */
  async deleteFolder(path: string): Promise<{ children: string[]; messageCount?: number; isLabel: boolean }> {
    validateFolderPath(path);
    const isLabel = /^Labels\//i.test(path);
    if (!/^(Folders|Labels)\//i.test(path)) {
      throw new Error(
        `delete_folder is restricted to the "Folders/" and "Labels/" namespaces to protect system mailboxes (got: "${path}").`,
      );
    }
    // Note: previously delete_folder also rejected `.`/`..` path segments as
    // defense-in-depth. That guard provided no actual security — IMAP treats
    // paths as opaque literal names, so `Folders/../INBOX` is just a different
    // mailbox name, not a parent reference. The check was blocking legitimate
    // cleanup of adversarial paths created by other IMAP clients (or older
    // versions of this MCP). The same guard remains on create/rename, which
    // is where keeping confusable paths out of existence actually matters.
    this.log(`[IMAP] Deleting folder ${path}`);
    const client = this.createClient();
    try {
      await client.connect();
      // Enumerate children before delete so we can surface the cascade in the
      // response. IMAP's DELETE on a parent folder with children is mostly
      // server-defined behavior; Proton Mail Bridge deletes the whole subtree.
      // Callers expecting "I'll delete the top-level one only" deserve to see
      // that they actually deleted N nested folders too.
      const children: string[] = [];
      try {
        const all = await client.list();
        for (const mb of all) {
          if (mb.path !== path && mb.path.startsWith(path + "/")) {
            children.push(mb.path);
          }
        }
      } catch {
        // Non-fatal — if LIST fails we'll still attempt the delete; we just
        // can't enumerate the cascade in the response.
      }
      // Count the messages about to be relocated (Folders) / untagged (Labels) so
      // the response can say where they went — deleting a non-empty folder on
      // Proton is non-destructive (contents move to All Mail), but a bare "Folder
      // deleted" reads as if the messages were destroyed. STATUS is read-only and
      // doesn't select the mailbox, so it's safe to run right before the delete.
      let messageCount: number | undefined;
      try {
        const st = await client.status(path, { messages: true });
        messageCount = st.messages ?? undefined;
      } catch {
        // Non-fatal — omit the count if STATUS isn't available for this path.
      }
      try {
        await client.mailboxDelete(path);
      } catch (err) {
        if (isMailboxMissingError(err)) {
          throw new Error(`Folder not found: ${path}`);
        }
        // Proton returns bare "Command failed" without mailboxMissing — probe
        // with status() to disambiguate, mirroring renameFolder.
        try {
          await client.status(path, { messages: true });
        } catch (probeErr) {
          if (isMailboxMissingError(probeErr)) {
            throw new Error(`Folder not found: ${path}`);
          }
        }
        throw err;
      }
      return { children, messageCount, isLabel };
    } finally {
      await client.logout().catch(() => {});
    }
  }

  async emptyFolder(folder: string, options: { dryRun?: boolean } = {}): Promise<{ expunged: number }> {
    validateFolderPath(folder);
    const dryRun = options.dryRun ?? false;
    this.log(`[IMAP] ${dryRun ? "Previewing empty of" : "Emptying"} folder ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const uids = (await client.search({ all: true }, { uid: true })) || [];
        if (uids.length === 0) {
          return { expunged: 0 };
        }
        // dryRun reports the count that WOULD be expunged without touching mail.
        if (dryRun) {
          return { expunged: uids.length };
        }
        // messageDelete performs STORE +\Deleted and EXPUNGE in one atomic step,
        // so partial failures don't leave messages flagged-but-not-expunged.
        await client.messageDelete(uids.join(","), { uid: true });
        return { expunged: uids.length };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Count messages matching criteria. Returns the SEARCH result length only;
   * no envelopes are fetched. Note: `hasAttachment` is stripped from criteria
   * because counting via the post-filter would require a full envelope scan,
   * defeating the purpose. Use search_messages for hasAttachment filtering.
   */
  async countMessages(folder: string, criteria: SearchCriteria): Promise<number> {
    validateFolderPath(folder);
    this.log(`[IMAP] Counting messages in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const query = this.buildSearchQuery({ ...criteria, hasAttachment: undefined });
        const result = await client.search(query, { uid: true });
        const count = result ? result.length : 0;
        if (count === 0) await this.assertSelectableFolder(folder);
        return count;
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Return aggregate stats for a folder: total/unread (via STATUS, free), plus
   * scanned-envelope aggregations (oldest/newest/totalBytes). Scans up to
   * scanLimit UIDs (highest = most recent). Always includes scanned/truncated
   * so callers can detect partial results.
   */
  async folderStats(folder: string, scanLimit: number): Promise<FolderStats> {
    validateFolderPath(folder);
    if (!Number.isInteger(scanLimit) || scanLimit < 1 || scanLimit > 20000) {
      throw new Error("scanLimit must be an integer between 1 and 20000");
    }
    this.log(`[IMAP] Computing folder stats for ${folder} (scanLimit=${scanLimit})`);
    const client = this.createClient();
    try {
      await client.connect();
      const status = await client.status(folder, { messages: true, unseen: true });
      const total = status.messages ?? 0;
      const unread = status.unseen ?? 0;

      const lock = await this.lockFolder(client, folder);
      try {
        const allUids = (await client.search({ all: true }, { uid: true })) || [];
        if (allUids.length === 0) await this.assertSelectableFolder(folder);
        const truncated = allUids.length > scanLimit;
        const scanUids = allUids.slice(-scanLimit); // highest UIDs (most recent)

        let oldest: Date | undefined;
        let newest: Date | undefined;
        let totalBytes = 0;
        let scanned = 0;

        if (scanUids.length > 0) {
          for await (const msg of client.fetch(
            scanUids.join(","),
            { uid: true, envelope: true, size: true },
            { uid: true },
          )) {
            scanned += 1;
            if (msg.size) totalBytes += msg.size;
            const d = msg.envelope?.date;
            if (d) {
              if (!oldest || d < oldest) oldest = d;
              if (!newest || d > newest) newest = d;
            }
          }
        }

        return {
          folder,
          total,
          unread,
          scanned,
          truncated,
          scanLimit,
          oldest: oldest?.toISOString(),
          newest: newest?.toISOString(),
          totalBytes,
        };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Return a frequency table of top senders. Buckets are keyed by lowercased
   * email address (falls back to lowercased name when address is empty).
   * Display name is carried forward from the latest seen entry that has one.
   * lastDate is the maximum date across all messages from that bucket.
   * Scans up to scanLimit UIDs (highest = most recent). `hasAttachment` is
   * stripped from criteria — use search_messages for that.
   */
  /**
   * Pick the most frequently observed display name for an address bucket.
   * Display names are attacker-controlled, so using the majority (rather than
   * first- or last-seen) prevents a single message with a spoofed `From` name
   * from poisoning the whole bucket label. Ties break lexicographically for
   * deterministic output.
   */
  private static pickDominantName(nameCounts: Map<string, number>): string | undefined {
    let best: string | undefined;
    let bestN = -1;
    for (const [name, n] of nameCounts) {
      if (n > bestN || (n === bestN && (best === undefined || name < best))) {
        best = name;
        bestN = n;
      }
    }
    return best;
  }

  async topSenders(
    folder: string,
    criteria: SearchCriteria,
    limit: number,
    scanLimit: number,
    options: { excludeSelf?: boolean; userAddress?: string } = {},
  ): Promise<TopSendersResult> {
    validateFolderPath(folder);
    if (!Number.isInteger(scanLimit) || scanLimit < 1 || scanLimit > 20000) {
      throw new Error("scanLimit must be an integer between 1 and 20000");
    }
    const userAddress = options.userAddress?.toLowerCase();
    const excludeSelf = Boolean(options.excludeSelf && userAddress);
    this.log(`[IMAP] Top senders for ${folder} (scanLimit=${scanLimit}, limit=${limit}, excludeSelf=${excludeSelf})`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const query = this.buildSearchQuery({ ...criteria, hasAttachment: undefined });
        const allUids = (await client.search(query, { uid: true })) || [];
        const truncated = allUids.length > scanLimit;
        const scanUids = allUids.slice(-scanLimit);

        // Track display-name frequency per address; the label uses the most
        // frequent name (pickDominantName) so a single spoofed `From` name can't
        // poison the bucket — the address is always shown alongside.
        const buckets = new Map<string, { nameCounts: Map<string, number>; count: number; lastDate?: Date }>();
        let scanned = 0;

        if (scanUids.length > 0) {
          for await (const msg of client.fetch(scanUids.join(","), { uid: true, envelope: true }, { uid: true })) {
            scanned += 1;
            const fromArr = msg.envelope?.from;
            if (!fromArr || fromArr.length === 0) continue;
            const f = fromArr[0];
            const addr = (f.address || "").toLowerCase();
            const key = addr || (f.name || "").toLowerCase();
            if (!key) continue;
            const existing = buckets.get(key);
            const d = msg.envelope?.date;
            if (existing) {
              existing.count += 1;
              if (f.name) existing.nameCounts.set(f.name, (existing.nameCounts.get(f.name) ?? 0) + 1);
              if (d && (!existing.lastDate || d > existing.lastDate)) existing.lastDate = d;
            } else {
              const nameCounts = new Map<string, number>();
              if (f.name) nameCounts.set(f.name, 1);
              buckets.set(key, { nameCounts, count: 1, lastDate: d });
            }
          }
        }

        let entries = [...buckets.entries()];
        if (excludeSelf && userAddress) {
          entries = entries.filter(([key]) => key !== userAddress);
        }
        const rows: TopSenderRow[] = entries
          .sort((a, b) => b[1].count - a[1].count)
          .slice(0, limit)
          .map(([key, v]) => {
            const name = ImapService.pickDominantName(v.nameCounts);
            const row: TopSenderRow = {
              from: name ? `${name} <${key}>` : key,
              count: v.count,
              lastDate: v.lastDate?.toISOString(),
            };
            if (userAddress) {
              row.direction = key === userAddress ? "self" : "received";
            }
            return row;
          });

        return { rows, scanned, truncated, scanLimit };
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  /**
   * Resolve the thread members the same way a live thread op would see them,
   * for use by `dryRun` previews. Mirrors `threadOp`'s folder-scoping logic:
   * acrossFolders=true walks DEFAULT_THREAD_FOLDERS; acrossFolders=false
   * scopes to the seed message's folder via firstFolderForMessageId.
   */
  async previewThread(messageId: string, acrossFolders: boolean): Promise<MessageSummary[]> {
    const folders = acrossFolders ? undefined : await this.firstFolderForMessageId(messageId);
    const members = await this.getThreadByMessageId(messageId, 1000, folders);
    // Reroute All Mail-only orphans regardless of acrossFolders. The original
    // gate was acrossFolders=true, but the acrossFolders=false path hits the
    // same bug: when the seed message lives in a user folder (Archive, etc.),
    // firstFolderForMessageId returns ["All Mail"] (since it walks
    // INBOX/Sent/All Mail) and the live mutation against All Mail UIDs
    // silently no-ops under Proton's label model. Same root cause as the
    // acrossFolders=true case the helper was originally added for. The
    // internal filter in rerouteAllMailOrphans (orphans only) is the actual
    // gate — calls with healthy INBOX/Sent rows short-circuit immediately.
    return await this.rerouteAllMailOrphans(members);
  }

  /**
   * Rewrite "All Mail"-only thread members to their actual storage folder.
   *
   * Background: the cross-folder walk surfaces messages by searching INBOX +
   * Sent + All Mail. When a thread member lives in a user folder (Archive,
   * Folders/X, etc.) and not in INBOX or Sent, the dedupe finds only the All
   * Mail copy and tags it `folder: "All Mail"` with `otherFolders: []`. A
   * subsequent move/delete against `All Mail` is a no-op under Proton's label
   * model — All Mail is a virtual aggregate, not a storage location, so
   * trying to remove the `\All` label fails silently. The thread op then
   * reports `0 affected, 1 notFound` even though the message exists.
   *
   * Fix: for each orphan member, scan the user mailbox list for a real folder
   * holding the Message-ID and rewrite the member's `folder` field to it.
   * Pay the extra IMAP round-trips only when the standard walk missed —
   * threads whose members were in INBOX/Sent never trigger this branch.
   *
   * Skipped folders: Labels/* (additive tags, not storage), Trash (deleting
   * from Trash isn't what the caller asked for), and All Mail itself.
   */
  private async rerouteAllMailOrphans<T extends MessageSummary & { folder?: string; otherFolders?: string[] }>(
    members: T[],
  ): Promise<T[]> {
    // A row counts as orphaned when (a) it lives in All Mail, (b) the dedupe
    // surfaced no other folder for the same Message-ID, AND (c) no other row
    // in the same result set carries the Message-ID in a real folder. The
    // last condition is defense-in-depth for callers/mocks that bypass the
    // dedupe path in `getThreadByMessageId` — the real implementation
    // collapses multi-folder copies into one row, but stubbed test paths can
    // return them as separate entries.
    const idsWithRealFolder = new Set<string>();
    for (const m of members) {
      if (m.messageId && m.folder && !/^all mail$/i.test(m.folder)) {
        idsWithRealFolder.add(m.messageId);
      }
    }
    const orphans = members.filter(
      (m) =>
        /^all mail$/i.test(m.folder ?? "") &&
        (!m.otherFolders || m.otherFolders.length === 0) &&
        m.messageId &&
        !idsWithRealFolder.has(m.messageId),
    );
    if (orphans.length === 0) return members;

    const client = this.createClient();
    try {
      await client.connect();
      const mailboxes = await client.list();
      // Candidate folders: user-storage paths only. Drop Labels/* (additive),
      // All Mail (virtual), and Trash (a delete against Trash means "expunge
      // from Trash" — which is what the caller would have explicitly asked
      // for if that's what they wanted).
      const candidates = mailboxes
        .map((mb) => mb.path)
        .filter((p) => !/^all mail$/i.test(p))
        .filter((p) => !/^trash$/i.test(p))
        .filter((p) => !/^labels\//i.test(p));

      for (const orphan of orphans) {
        const mid = orphan.messageId!;
        let found: string | undefined;
        let foundUid: number | undefined;
        for (const folder of candidates) {
          try {
            const lock = await this.lockFolder(client, folder);
            try {
              const result = await client.search({ header: { "Message-ID": mid } }, { uid: true });
              if (result && result.length > 0) {
                found = folder;
                // Capture the UID *in this real folder* too. The orphan still
                // carries its All Mail UID, which is invalid in `found` — leaving
                // it would make the dry-run preview pair (e.g.) "Archive / UID
                // 189" where 189 is the All Mail UID, not Archive's. The live
                // path re-resolves by Message-ID so it was unaffected, but an
                // agent trusting the previewed pairing in a single-message tool
                // would target the wrong message. Take the highest UID (newest)
                // when a Message-ID somehow matches more than one row.
                foundUid = result[result.length - 1];
                break;
              }
            } finally {
              lock.release();
            }
          } catch {
            // Folder may have disappeared between list() and lockFolder() —
            // skip and keep looking.
          }
        }
        if (found) {
          orphan.folder = found;
          if (foundUid !== undefined) orphan.uid = foundUid;
        }
      }
      return members;
    } finally {
      await client.logout().catch(() => {});
    }
  }

  async moveThread(messageId: string, destination: string, acrossFolders: boolean): Promise<ThreadOpResult> {
    return this.threadOp(messageId, acrossFolders, async (folder, uids) => {
      if (folder === destination) {
        // Skip self-move to avoid IMAP errors when one of the walked folders == destination.
        return { folder, affected: 0, notFound: [] };
      }
      const r = await this.bulkMove(folder, uids, destination);
      return { folder, affected: r.moved, notFound: r.notFound };
    });
  }

  async deleteThread(messageId: string, permanent: boolean, acrossFolders: boolean): Promise<ThreadOpResult> {
    return this.threadOp(messageId, acrossFolders, async (folder, uids) => {
      const r = await this.bulkDelete(folder, uids, permanent);
      return { folder, affected: r.deleted, notFound: r.notFound };
    });
  }

  async flagThread(
    messageId: string,
    add: string[],
    remove: string[],
    acrossFolders: boolean,
  ): Promise<ThreadOpResult> {
    return this.threadOp(messageId, acrossFolders, async (folder, uids) => {
      const r = await this.bulkUpdateFlags(folder, uids, add, remove);
      return { folder, affected: r.affected, notFound: r.notFound };
    });
  }

  /**
   * Re-resolve current UIDs in `folder` for any of the given Message-IDs. Used
   * by threadOp to sidestep stale-UID accounting on Proton's label model: when
   * a thread member is moved out of one folder, Proton may cascade the change
   * into another folder's view (e.g., moving from INBOX to Trash strips the
   * `All Mail` label too, since All Mail excludes Trash). If we operated on
   * upfront-snapshotted UIDs we'd see them as `notFound` even though the
   * underlying message did move. Re-resolving by Message-ID at the moment we
   * touch each folder produces accurate per-folder accounting.
   */
  private async resolveUidsByMessageIds(folder: string, messageIds: string[]): Promise<number[]> {
    if (messageIds.length === 0) return [];
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const uids = new Set<number>();
        for (const mid of messageIds) {
          if (!mid) continue;
          const found = await client.search({ header: { "Message-ID": mid } }, { uid: true });
          if (found && found.length > 0) for (const u of found) uids.add(u);
        }
        return [...uids];
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  private async threadOp(
    messageId: string,
    acrossFolders: boolean,
    op: (folder: string, uids: number[]) => Promise<PerFolderResult>,
  ): Promise<ThreadOpResult> {
    // Gather the thread membership upfront — both UIDs (for the dry-run path,
    // which calls previewThread separately) and Message-IDs (for per-folder
    // re-resolution at mutation time). Each member carries its folder tag and
    // Message-ID via MessageSummary.
    const members = await this.previewThread(messageId, acrossFolders);

    // Group by folder, deduping Message-IDs since labeled-self threads can
    // surface the same Message-ID twice in the same folder under different UIDs.
    const messageIdsByFolder = new Map<string, Set<string>>();
    for (const m of members) {
      const folder = (m as MessageSummary & { folder?: string }).folder ?? "INBOX";
      const mid = m.messageId ?? "";
      if (!mid) continue;
      const set = messageIdsByFolder.get(folder) ?? new Set<string>();
      set.add(mid);
      messageIdsByFolder.set(folder, set);
    }

    const perFolder: PerFolderResult[] = [];
    let total = 0;
    for (const [folder, messageIds] of messageIdsByFolder.entries()) {
      try {
        // Re-resolve UIDs in this folder *now*, after any prior-folder operations
        // have cascaded through Proton's label model. Empty result means the
        // messages this folder claimed to contain are already gone — report as
        // a clean 0/0, not as spurious notFound entries.
        const currentUids = await this.resolveUidsByMessageIds(folder, [...messageIds]);
        if (currentUids.length === 0) {
          perFolder.push({ folder, affected: 0, notFound: [] });
          continue;
        }
        const r = await op(folder, currentUids);
        perFolder.push(r);
        total += r.affected;
      } catch (err) {
        perFolder.push({
          folder,
          affected: 0,
          notFound: [],
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }
    return { perFolder, total };
  }

  private async firstFolderForMessageId(messageId: string): Promise<string[] | undefined> {
    // Walk DEFAULT_THREAD_FOLDERS, return the first one where the message-id is found, as a single-element list.
    for (const folder of DEFAULT_THREAD_FOLDERS) {
      const uid = await this.findByMessageId(folder, messageId).catch(() => undefined);
      if (uid !== undefined) return [folder];
    }
    return undefined; // caller will use DEFAULT_THREAD_FOLDERS as fallback
  }

  /**
   * Attach a ~200-char body preview (snippet) to each MessageSummary in place.
   * Downloads text/plain (preferred) or stripped text/html for each message.
   * On any per-message failure, sets snippet to "" rather than propagating the error.
   * Must be called while the mailbox lock is held (uses the same client connection).
   */
  private async attachSnippets(client: ImapFlow, messages: MessageSummary[]): Promise<void> {
    for (const msg of messages) {
      try {
        const meta = await client.fetchOne(String(msg.uid), { uid: true, bodyStructure: true }, { uid: true });
        if (!meta || !meta.bodyStructure) {
          msg.snippet = "";
          continue;
        }
        const textPart = findPartNumber(meta.bodyStructure, "text/plain");
        const htmlPart = findPartNumber(meta.bodyStructure, "text/html");
        let raw = "";
        if (textPart) {
          raw = await this.downloadPart(client, msg.uid, textPart);
        } else if (htmlPart) {
          raw = stripHtml(await this.downloadPart(client, msg.uid, htmlPart));
        }
        const collapsed = raw.replace(/\s+/g, " ").trim();
        msg.snippet = collapsed.length > 200 ? collapsed.slice(0, 200) + "…" : collapsed;
      } catch {
        msg.snippet = "";
      }
    }
  }

  private toSummary(msg: FetchMessageObject): MessageSummary {
    return {
      uid: msg.uid,
      subject: msg.envelope?.subject || "",
      from: formatAddress(msg.envelope?.from),
      to: formatAddress(msg.envelope?.to),
      date: msg.envelope?.date?.toISOString() || "",
      flags: msg.flags ? [...msg.flags] : [],
      messageId: msg.envelope?.messageId || "",
    };
  }
}
