import { ImapFlow } from "imapflow";
import type {
  FetchMessageObject,
  MessageAddressObject,
  MessageStructureObject,
  SearchObject,
} from "imapflow";

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
  text: string;
  html: string;
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
  return addr
    .map((a) => (a.name ? `${a.name} <${a.address}>` : a.address || ""))
    .join(", ");
}

/**
 * IMAP service for reading emails via Proton Mail Bridge.
 * Each operation creates a fresh connection to avoid stale/idle connection issues.
 */
export class ImapService {
  private config: ImapConfig;
  private debug: boolean;

  constructor(config: ImapConfig) {
    this.config = config;
    this.debug = config.debug || false;
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
  async listMessages(folder: string, limit: number): Promise<MessageSummary[]> {
    this.log(`[IMAP] Listing messages from ${folder}, limit ${limit}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        const status = await client.status(folder, { messages: true });
        const total = status.messages ?? 0;
        if (total === 0) return [];

        const start = Math.max(1, total - limit + 1);
        const range = `${start}:*`;

        const messages: MessageSummary[] = [];
        for await (const msg of client.fetch(range, {
          uid: true,
          envelope: true,
          flags: true,
        })) {
          messages.push(this.toSummary(msg));
        }

        // Return newest first
        messages.reverse();
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
   */
  async readMessage(folder: string, uid: number): Promise<MessageDetail> {
    this.log(`[IMAP] Reading message UID ${uid} from ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        const msg = await client.fetchOne(
          String(uid),
          {
            uid: true,
            envelope: true,
            flags: true,
            bodyStructure: true,
            bodyParts: ["TEXT", "1"],
          },
          { uid: true },
        );

        if (!msg) {
          throw new Error(`Message UID ${uid} not found in ${folder}`);
        }

        const textPart = msg.bodyParts?.get("TEXT") || msg.bodyParts?.get("1");
        const textContent = textPart ? textPart.toString("utf-8") : "";

        let text = "";
        let html = "";
        if (msg.bodyStructure && ImapService.hasHtmlPart(msg.bodyStructure)) {
          html = textContent;
        } else {
          text = textContent;
        }

        const attachments = msg.bodyStructure
          ? ImapService.extractAttachments(msg.bodyStructure)
          : [];

        return {
          uid: msg.uid,
          subject: msg.envelope?.subject || "",
          from: formatAddress(msg.envelope?.from),
          to: formatAddress(msg.envelope?.to),
          cc: formatAddress(msg.envelope?.cc),
          date: msg.envelope?.date?.toISOString() || "",
          messageId: msg.envelope?.messageId || "",
          flags: msg.flags ? [...msg.flags] : [],
          text,
          html,
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
        for await (const msg of client.fetch(uidRange, {
          uid: true,
          envelope: true,
          flags: true,
        }, { uid: true })) {
          messages.push(this.toSummary(msg));
        }

        messages.reverse();
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
        filename:
          structure.dispositionParameters?.filename ||
          structure.parameters?.name ||
          "unnamed",
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
    this.log(`[IMAP] Downloading attachment part ${partNumber} from UID ${uid} in ${folder}`);
    const client = this.createClient();
    try {
      await client.connect();
      const lock = await client.getMailboxLock(folder);
      try {
        const { meta, content } = await client.download(String(uid), partNumber, { uid: true });
        const chunks: Buffer[] = [];
        for await (const chunk of content) {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
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

  private static hasHtmlPart(structure: MessageStructureObject): boolean {
    if (structure.type === "text/html") return true;
    if (structure.childNodes) {
      return structure.childNodes.some((child) => ImapService.hasHtmlPart(child));
    }
    return false;
  }

  /**
   * Move a message to a different folder.
   */
  async moveMessage(folder: string, uid: number, destination: string): Promise<boolean> {
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
