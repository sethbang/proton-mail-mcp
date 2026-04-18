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
}

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
 * Returns the first match (depth-first).
 *
 * Single-part messages have no `part` field on the root node — IMAP
 * implicitly treats those as part "1".
 */
function findPartNumber(structure: MessageStructureObject, targetType: string): string | undefined {
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
 */
function stripHtml(html: string): string {
  return convert(html, {
    wordwrap: false,
    selectors: [
      { selector: "a", options: { hideLinkHrefIfSameAsText: true } },
      { selector: "img", format: "skip" },
    ],
  });
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
      const lock = await client.getMailboxLock(folder);
      try {
        let result: number[] | false;
        if (beforeUid) {
          result = await client.search({ uid: `1:${beforeUid - 1}` }, { uid: true });
        } else {
          result = await client.search({ all: true }, { uid: true });
        }
        if (!result || result.length === 0) return [];

        return this.fetchSortAndLimit(client, result, limit);
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
    options?: { preferHtml?: boolean; maxBodyLength?: number },
  ): Promise<MessageDetail> {
    validateFolderPath(folder);
    const preferHtml = options?.preferHtml ?? false;
    const maxLen = options?.maxBodyLength ?? DEFAULT_MAX_BODY_LENGTH;
    this.log(`[IMAP] Reading message UID ${uid} from ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        // Step 1: fetch metadata and structure (no body content yet)
        const msg = await client.fetchOne(
          String(uid),
          {
            uid: true,
            envelope: true,
            flags: true,
            bodyStructure: true,
          },
          { uid: true },
        );

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
          body = stripHtml(rawHtml);
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
      const lock = await client.getMailboxLock(folder);
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

        return this.fetchSortAndLimit(client, result, limit);
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
      const lock = await client.getMailboxLock(folder);
      try {
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
   * Move a message to a different folder.
   */
  async moveMessage(folder: string, uid: number, destination: string): Promise<MoveResult> {
    validateFolderPath(folder);
    validateFolderPath(destination);
    this.log(`[IMAP] Moving UID ${uid} from ${folder} to ${destination}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        const result = await client.messageMove(String(uid), destination, { uid: true });
        if (result === false) {
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
      const lock = await client.getMailboxLock(folder);
      try {
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
      const lock = await client.getMailboxLock(folder);
      try {
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
      const lock = await client.getMailboxLock(folder);
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
      const lock = await client.getMailboxLock(folder);
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
   * Searches within a single folder.
   */
  async getThread(folder: string, uid: number, limit: number): Promise<MessageSummary[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Getting thread for UID ${uid} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        // Step 1: fetch the seed message's messageId and headers
        const seed = await client.fetchOne(
          String(uid),
          { uid: true, envelope: true, headers: ["references"] },
          { uid: true },
        );
        if (!seed) throw new Error(`Message UID ${uid} not found in ${folder}`);

        const seedMessageId = seed.envelope?.messageId || "";
        if (!seedMessageId) throw new Error(`Message UID ${uid} has no Message-ID`);

        // Parse References header from the raw headers buffer
        const headersStr = seed.headers ? seed.headers.toString("utf-8") : "";
        const referencesMatch = headersStr.match(/^References:\s*(.+(?:\r?\n[ \t]+.+)*)/im);
        const referencesRaw = referencesMatch ? referencesMatch[1].replace(/\r?\n[ \t]+/g, " ") : "";
        const inReplyTo = seed.envelope?.inReplyTo || "";
        const knownIds = new Set<string>();
        knownIds.add(seedMessageId);

        // Extract message-IDs from References header (space or newline separated)
        const refMatches = referencesRaw.match(/<[^>]+>/g);
        if (refMatches) {
          for (const ref of refMatches) knownIds.add(ref);
        }
        if (inReplyTo) knownIds.add(inReplyTo);

        // Step 2: search for all messages that reference any known ID, or are referenced
        const allUids = new Set<number>();
        allUids.add(uid);

        for (const msgId of knownIds) {
          // Find messages with this Message-ID
          const byId = await client.search({ header: { "Message-ID": msgId } }, { uid: true });
          if (byId && byId.length > 0) {
            for (const u of byId) allUids.add(u);
          }
          // Find messages that reference this Message-ID
          const byRef = await client.search({ header: { References: msgId } }, { uid: true });
          if (byRef && byRef.length > 0) {
            for (const u of byRef) allUids.add(u);
          }
          // Also check In-Reply-To
          const byReply = await client.search({ header: { "In-Reply-To": msgId } }, { uid: true });
          if (byReply && byReply.length > 0) {
            for (const u of byReply) allUids.add(u);
          }

          if (allUids.size >= limit) break;
        }

        // Step 3: fetch envelopes for all found UIDs
        const uidList = [...allUids].slice(0, limit);
        const fetchRange = uidList.join(",");
        const messages: MessageSummary[] = [];
        for await (const msg of client.fetch(fetchRange, { uid: true, envelope: true, flags: true }, { uid: true })) {
          messages.push(this.toSummary(msg));
        }

        // Sort chronologically (oldest first for thread reading)
        messages.sort((a, b) => {
          const da = a.date ? new Date(a.date).getTime() : 0;
          const db = b.date ? new Date(b.date).getTime() : 0;
          return da !== db ? da - db : a.uid - b.uid;
        });
        return messages;
      } finally {
        lock.release();
      }
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
      const result = await client.append(folder, rawMessage, ["\\Draft", "\\Seen"]);
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
