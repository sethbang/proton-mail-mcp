import { ImapFlow } from "imapflow";
import type {
  FetchMessageObject,
  MessageAddressObject,
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
            bodyParts: ["TEXT", "1"],
          },
          { uid: true },
        );

        if (!msg) {
          throw new Error(`Message UID ${uid} not found in ${folder}`);
        }

        const textPart = msg.bodyParts?.get("TEXT") || msg.bodyParts?.get("1");
        const textContent = textPart ? textPart.toString("utf-8") : "";

        // Try to separate HTML from plain text based on content
        let text = "";
        let html = "";
        if (textContent.includes("<html") || textContent.includes("<body") || textContent.includes("<div")) {
          html = textContent;
        } else {
          text = textContent;
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
          text,
          html,
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
