import { ImapFlow } from "imapflow";
import type { FetchMessageObject, MessageAddressObject, MessageStructureObject, SearchObject } from "imapflow";
import { convert } from "html-to-text";
import { validateFolderPath, validatePartNumber } from "./validation.js";

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
}

export interface AttachmentMeta {
  partNumber: string;
  filename: string;
  contentType: string;
  size: number;
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

export interface FolderInfo {
  path: string;
  name: string;
  specialUse: string;
  messages: number;
  unseen: number;
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
      }));
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

  async listMessages(folder: string, limit: number, beforeUid?: number): Promise<MessageSummary[]> {
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
        if (!result || result.length === 0) return [];

        return await this.fetchSortAndLimit(client, result, limit);
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
   * Search messages in a folder by criteria.
   */
  async searchMessages(
    folder: string,
    criteria: {
      from?: string;
      to?: string;
      subject?: string;
      body?: string;
      since?: string;
      before?: string;
      seen?: boolean;
      flagged?: boolean;
    },
    limit: number,
  ): Promise<MessageSummary[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Searching in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const query: SearchObject = {};
        if (criteria.from) query.from = criteria.from;
        if (criteria.to) query.to = criteria.to;
        if (criteria.subject) query.subject = criteria.subject;
        if (criteria.body) query.body = criteria.body;
        if (criteria.since) query.since = criteria.since;
        if (criteria.before) query.before = criteria.before;
        if (criteria.seen !== undefined) query.seen = criteria.seen;
        if (criteria.flagged !== undefined) query.flagged = criteria.flagged;

        const result = await client.search(query, { uid: true });
        if (!result || result.length === 0) return [];

        return await this.fetchSortAndLimit(client, result, limit);
      } finally {
        lock.release();
      }
    } finally {
      await client.logout().catch(() => {});
    }
  }

  private static extractAttachments(structure: MessageStructureObject): AttachmentMeta[] {
    const attachments: AttachmentMeta[] = [];

    if (
      structure.disposition === "attachment" ||
      (structure.disposition === "inline" && structure.type && !structure.type.startsWith("text/"))
    ) {
      attachments.push({
        partNumber: structure.part || "",
        filename: structure.dispositionParameters?.filename || structure.parameters?.name || "unnamed",
        contentType: structure.type || "application/octet-stream",
        size: structure.size || 0,
      });
    }

    if (structure.childNodes) {
      for (const child of structure.childNodes) {
        attachments.push(...ImapService.extractAttachments(child));
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
        if (!knownAttachments.some((a) => a.partNumber === partNumber)) {
          const known = knownAttachments.map((a) => a.partNumber).join(", ") || "(none)";
          throw new Error(`Part ${partNumber} not found on UID ${uid} in ${folder}; known parts: [${known}]`);
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
   * Add and/or remove flags on a message.
   */
  async updateFlags(folder: string, uid: number, flagsToAdd: string[], flagsToRemove: string[]): Promise<boolean> {
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
        return true;
      } finally {
        lock.release();
      }
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
  async markAllRead(folder: string, olderThan?: string): Promise<number> {
    validateFolderPath(folder);
    this.log(`[IMAP] Marking all unread as read in ${folder}${olderThan ? ` before ${olderThan}` : ""}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await this.lockFolder(client, folder);
      try {
        const query: SearchObject = { seen: false };
        if (olderThan) query.before = olderThan;

        const result = await client.search(query, { uid: true });
        if (!result || result.length === 0) return 0;

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

      // De-duplicate by Message-ID would be ideal but our summaries don't carry it;
      // UIDs are per-folder so the { folder, uid } tuple is unique.
      merged.sort((a, b) => {
        const da = a.date ? new Date(a.date).getTime() : 0;
        const db = b.date ? new Date(b.date).getTime() : 0;
        return da !== db ? da - db : a.uid - b.uid;
      });
      return merged.slice(0, limit);
    } finally {
      await client.logout().catch(() => {});
    }
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

  private toSummary(msg: FetchMessageObject): MessageSummary {
    return {
      uid: msg.uid,
      subject: msg.envelope?.subject || "",
      from: formatAddress(msg.envelope?.from),
      to: formatAddress(msg.envelope?.to),
      date: msg.envelope?.date?.toISOString() || "",
      flags: msg.flags ? [...msg.flags] : [],
    };
  }
}
