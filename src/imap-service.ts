import { ImapFlow } from "imapflow";
import type { FetchMessageObject, MessageAddressObject, MessageStructureObject, SearchObject } from "imapflow";
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
  attachments: AttachmentMeta[];
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
 * Strip HTML tags and decode common HTML entities to produce readable plain text.
 */
function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/li>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
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
   * List recent messages from a folder.
   */
  async listMessages(folder: string, limit: number, beforeUid?: number): Promise<MessageSummary[]> {
    validateFolderPath(folder);
    this.log(`[IMAP] Listing messages from ${folder}, limit ${limit}${beforeUid ? `, beforeUid ${beforeUid}` : ""}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        const messages: MessageSummary[] = [];

        if (beforeUid) {
          const uidRange = `1:${beforeUid - 1}`;
          const result = await client.search({ uid: uidRange }, { uid: true });
          if (!result || result.length === 0) return [];

          const selectedUids = result.slice(-limit);
          const fetchRange = selectedUids.join(",");
          for await (const msg of client.fetch(
            fetchRange,
            {
              uid: true,
              envelope: true,
              flags: true,
            },
            { uid: true },
          )) {
            messages.push(this.toSummary(msg));
          }
        } else {
          const status = await client.status(folder, { messages: true });
          const total = status.messages ?? 0;
          if (total === 0) return [];

          const start = Math.max(1, total - limit + 1);
          const range = `${start}:*`;
          for await (const msg of client.fetch(range, {
            uid: true,
            envelope: true,
            flags: true,
          })) {
            messages.push(this.toSummary(msg));
          }
        }

        // Sort newest first by date, falling back to UID descending
        messages.sort((a, b) => {
          const da = a.date ? new Date(a.date).getTime() : 0;
          const db = b.date ? new Date(b.date).getTime() : 0;
          return da !== db ? db - da : b.uid - a.uid;
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
        if (body.length > maxLen) {
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

        // Take the most recent UIDs up to the limit
        const selectedUids = result.slice(-limit);
        const uidRange = selectedUids.join(",");

        const messages: MessageSummary[] = [];
        for await (const msg of client.fetch(
          uidRange,
          {
            uid: true,
            envelope: true,
            flags: true,
          },
          { uid: true },
        )) {
          messages.push(this.toSummary(msg));
        }

        // Sort newest first by date, falling back to UID descending
        messages.sort((a, b) => {
          const da = a.date ? new Date(a.date).getTime() : 0;
          const db = b.date ? new Date(b.date).getTime() : 0;
          return da !== db ? db - da : b.uid - a.uid;
        });
        return messages;
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
  async moveMessage(folder: string, uid: number, destination: string): Promise<boolean> {
    validateFolderPath(folder);
    validateFolderPath(destination);
    this.log(`[IMAP] Moving UID ${uid} from ${folder} to ${destination}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        const result = await client.messageMove(String(uid), destination, { uid: true });
        return result !== false;
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
