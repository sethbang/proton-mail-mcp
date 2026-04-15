#!/usr/bin/env node

/**
 * Proton Mail MCP Server
 *
 * This MCP server provides email sending and reading functionality
 * using Proton Mail's SMTP service and Proton Mail Bridge IMAP.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { EmailService, EmailConfig } from "./email-service.js";
import { ImapService, ImapConfig } from "./imap-service.js";
import { isValidEmailAddress, validateImapFlag, sanitizeErrorMessage } from "./validation.js";

// Get environment variables for SMTP configuration
const PROTONMAIL_USERNAME = process.env.PROTONMAIL_USERNAME;
const PROTONMAIL_PASSWORD = process.env.PROTONMAIL_PASSWORD;
const PROTONMAIL_HOST = process.env.PROTONMAIL_HOST || "smtp.protonmail.ch";
const rawPort = parseInt(process.env.PROTONMAIL_PORT || "587", 10);
const PROTONMAIL_PORT = Number.isNaN(rawPort) ? 587 : rawPort;
const PROTONMAIL_SECURE = process.env.PROTONMAIL_SECURE === "true";
const DEBUG = process.env.DEBUG === "true";

// Get environment variables for IMAP configuration (Proton Mail Bridge)
const IMAP_HOST = process.env.IMAP_HOST || "127.0.0.1";
const rawImapPort = parseInt(process.env.IMAP_PORT || "1143", 10);
const IMAP_PORT = Number.isNaN(rawImapPort) ? 1143 : rawImapPort;
const IMAP_SECURE = process.env.IMAP_SECURE === "true";
const IMAP_USERNAME = process.env.IMAP_USERNAME || PROTONMAIL_USERNAME;
const IMAP_PASSWORD = process.env.IMAP_PASSWORD || PROTONMAIL_PASSWORD;

// Validate required environment variables
if (!PROTONMAIL_USERNAME || !PROTONMAIL_PASSWORD) {
  console.error(
    "[Error] Missing required environment variables: PROTONMAIL_USERNAME and PROTONMAIL_PASSWORD must be set",
  );
  process.exit(1);
}

if (!IMAP_USERNAME || !IMAP_PASSWORD) {
  console.error(
    "[Error] Missing IMAP credentials: set IMAP_USERNAME/IMAP_PASSWORD or PROTONMAIL_USERNAME/PROTONMAIL_PASSWORD as fallback",
  );
  process.exit(1);
}

// Helper function for debug logging
function debugLog(message: string): void {
  if (DEBUG) {
    console.error(message);
  }
}

/**
 * Produce a safe error message for MCP clients. Delegates to validation.ts
 * which categorizes errors without exposing raw messages that may contain credentials.
 */
function sanitizeError(error: unknown): string {
  return sanitizeErrorMessage(error);
}

// Create email service configuration
const emailConfig: EmailConfig = {
  host: PROTONMAIL_HOST,
  port: PROTONMAIL_PORT,
  secure: PROTONMAIL_SECURE,
  auth: {
    user: PROTONMAIL_USERNAME,
    pass: PROTONMAIL_PASSWORD,
  },
  debug: DEBUG,
  connectionTimeout: 30_000,
  greetingTimeout: 30_000,
  socketTimeout: 60_000,
};

// Create IMAP service configuration
const imapConfig: ImapConfig = {
  host: IMAP_HOST,
  port: IMAP_PORT,
  secure: IMAP_SECURE,
  auth: {
    user: IMAP_USERNAME!,
    pass: IMAP_PASSWORD!,
  },
  debug: DEBUG,
  connectionTimeout: 30_000,
};

// Initialize services
const emailService = new EmailService(emailConfig);
const imapService = new ImapService(imapConfig);

/**
 * Create an MCP server with capabilities for tools
 */
const server = new McpServer({
  name: "proton-mail-mcp",
  version: "0.1.0",
});

// Validate comma-separated email addresses, rejecting dangerous characters
function validateAddresses(value: string): boolean {
  return value.split(",").every((addr) => isValidEmailAddress(addr.trim()));
}

// ─── Rate Limiter ───────────────────────────────────────────────────────────

const sendRateLimiter = {
  timestamps: [] as number[],
  maxPerMinute: 10,
  check(): boolean {
    const now = Date.now();
    this.timestamps = this.timestamps.filter((t) => now - t < 60_000);
    if (this.timestamps.length >= this.maxPerMinute) return false;
    this.timestamps.push(now);
    return true;
  },
};

// ─── SMTP Tools ─────────────────────────────────────────────────────────────

server.registerTool(
  "send_email",
  {
    description: "Send an email using Proton Mail SMTP",
    inputSchema: {
      to: z
        .string()
        .min(1, "Recipient is required")
        .max(10_000, "To field too long")
        .refine(validateAddresses, "Each 'to' address must be a valid email")
        .describe("Recipient email address(es). Multiple addresses can be separated by commas."),
      subject: z
        .string()
        .min(1, "Subject is required")
        .max(998, "Subject exceeds RFC 5322 line length limit")
        .describe("Email subject line"),
      body: z
        .string()
        .min(1, "Body is required")
        .max(500_000, "Body too large")
        .describe("Email body content (can be plain text or HTML)"),
      isHtml: z.boolean().optional().default(false).describe("Whether the body contains HTML content"),
      cc: z
        .string()
        .max(10_000, "CC field too long")
        .refine(validateAddresses, "Each CC address must be a valid email")
        .optional()
        .describe("CC recipient(s), separated by commas"),
      bcc: z
        .string()
        .max(10_000, "BCC field too long")
        .refine(validateAddresses, "Each BCC address must be a valid email")
        .optional()
        .describe("BCC recipient(s), separated by commas"),
      replyTo: z
        .string()
        .max(10_000, "Reply-To field too long")
        .refine(validateAddresses, "Reply-To must be a valid email")
        .optional()
        .describe("Reply-To email address"),
      fromName: z.string().max(200, "From name too long").optional().describe("Display name for the From field"),
      attachments: z
        .array(
          z.object({
            filename: z.string().min(1).describe("Attachment filename"),
            content: z.string().min(1).describe("Base64-encoded file content"),
            contentType: z.string().min(1).describe("MIME type (e.g. application/pdf, image/png)"),
          }),
        )
        .optional()
        .describe("File attachments (base64-encoded content)"),
    },
  },
  async ({ to, subject, body, isHtml, cc, bcc, replyTo, fromName, attachments }) => {
    debugLog(`[Tool] Executing tool: send_email`);

    if (!sendRateLimiter.check()) {
      return {
        content: [{ type: "text" as const, text: "Rate limit exceeded. Maximum 10 emails per minute." }],
        isError: true,
      };
    }

    try {
      await emailService.sendEmail({
        to,
        subject,
        body,
        isHtml,
        cc,
        bcc,
        replyTo,
        fromName,
        attachments,
      });

      return {
        content: [
          {
            type: "text" as const,
            text: `Email sent successfully to ${to}${cc ? ` with CC to ${cc}` : ""}${bcc ? ` and BCC to ${bcc}` : ""}.`,
          },
        ],
      };
    } catch (error) {
      console.error(`[Error] Failed to send email: ${error instanceof Error ? error.message : String(error)}`);

      return {
        content: [
          {
            type: "text" as const,
            text: `Failed to send email: ${sanitizeError(error)}`,
          },
        ],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "reply_email",
  {
    description:
      "Reply to an email message. Reads the original message and sends a reply with proper threading headers (In-Reply-To, References).",
    inputSchema: {
      uid: z.number().int().min(1).describe("UID of the message to reply to"),
      folder: z
        .string()
        .optional()
        .default("INBOX")
        .describe("Folder containing the original message (default: INBOX)"),
      body: z.string().min(1).max(500_000).describe("Reply body content"),
      isHtml: z.boolean().optional().default(false).describe("Whether the body contains HTML content"),
      cc: z
        .string()
        .max(10_000)
        .refine(validateAddresses, "Each CC address must be a valid email")
        .optional()
        .describe("Additional CC recipients, separated by commas"),
      bcc: z
        .string()
        .max(10_000)
        .refine(validateAddresses, "Each BCC address must be a valid email")
        .optional()
        .describe("BCC recipients, separated by commas"),
      replyAll: z
        .boolean()
        .optional()
        .default(false)
        .describe("Reply to all recipients (sender + TO + CC) instead of just sender"),
    },
  },
  async ({ uid, folder, body, isHtml, cc, bcc, replyAll }) => {
    debugLog(`[Tool] Executing tool: reply_email (uid=${uid}, folder=${folder}, replyAll=${replyAll})`);

    if (!sendRateLimiter.check()) {
      return {
        content: [{ type: "text" as const, text: "Rate limit exceeded. Maximum 10 emails per minute." }],
        isError: true,
      };
    }

    try {
      const original = await imapService.readMessage(folder, uid);

      // Build threading headers
      const inReplyTo = original.messageId;
      const references = original.messageId;

      // Build subject
      const subject = /^re:/i.test(original.subject) ? original.subject : `Re: ${original.subject}`;

      // Build recipients
      let to = original.from;
      let replyCC = cc || "";

      if (replyAll) {
        // Collect original TO and CC, excluding our own address
        const allRecipients = [original.to, original.cc]
          .filter(Boolean)
          .join(", ")
          .split(",")
          .map((a) => a.trim())
          .filter((a) => a && !a.includes(emailConfig.auth.user));

        if (allRecipients.length > 0) {
          replyCC = replyCC ? `${replyCC}, ${allRecipients.join(", ")}` : allRecipients.join(", ");
        }
      }

      await emailService.sendEmail({
        to,
        subject,
        body,
        isHtml,
        cc: replyCC || undefined,
        bcc,
        inReplyTo,
        references,
      });

      return {
        content: [
          {
            type: "text" as const,
            text: `Reply sent to ${to}${replyCC ? ` with CC to ${replyCC}` : ""}.`,
          },
        ],
      };
    } catch (error) {
      console.error(`[Error] Failed to reply: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to reply: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "forward_email",
  {
    description:
      "Forward an email message. Reads the original message and sends it to new recipients with proper threading headers.",
    inputSchema: {
      uid: z.number().int().min(1).describe("UID of the message to forward"),
      folder: z
        .string()
        .optional()
        .default("INBOX")
        .describe("Folder containing the original message (default: INBOX)"),
      to: z
        .string()
        .min(1, "Recipient is required")
        .max(10_000)
        .refine(validateAddresses, "Each 'to' address must be a valid email")
        .describe("Recipient email address(es), separated by commas"),
      body: z
        .string()
        .max(500_000)
        .optional()
        .default("")
        .describe("Optional message to prepend above the forwarded content"),
      isHtml: z.boolean().optional().default(false).describe("Whether the body contains HTML content"),
      cc: z
        .string()
        .max(10_000)
        .refine(validateAddresses, "Each CC address must be a valid email")
        .optional()
        .describe("CC recipients, separated by commas"),
      bcc: z
        .string()
        .max(10_000)
        .refine(validateAddresses, "Each BCC address must be a valid email")
        .optional()
        .describe("BCC recipients, separated by commas"),
    },
  },
  async ({ uid, folder, to, body, isHtml, cc, bcc }) => {
    debugLog(`[Tool] Executing tool: forward_email (uid=${uid}, folder=${folder}, to=${to})`);

    if (!sendRateLimiter.check()) {
      return {
        content: [{ type: "text" as const, text: "Rate limit exceeded. Maximum 10 emails per minute." }],
        isError: true,
      };
    }

    try {
      const original = await imapService.readMessage(folder, uid);

      const subject = /^fwd:/i.test(original.subject) ? original.subject : `Fwd: ${original.subject}`;

      const originalContent = original.text || original.html || "(no content)";
      const separator = "\n\n---------- Forwarded message ----------\n";
      const originalHeaders = `From: ${original.from}\nDate: ${original.date}\nSubject: ${original.subject}\nTo: ${original.to}\n\n`;
      const fullBody = body
        ? `${body}${separator}${originalHeaders}${originalContent}`
        : `${separator}${originalHeaders}${originalContent}`;

      await emailService.sendEmail({
        to,
        subject,
        body: fullBody,
        isHtml,
        cc,
        bcc,
        inReplyTo: original.messageId,
        references: original.messageId,
      });

      return {
        content: [
          {
            type: "text" as const,
            text: `Message forwarded to ${to}${cc ? ` with CC to ${cc}` : ""}.`,
          },
        ],
      };
    } catch (error) {
      console.error(`[Error] Failed to forward: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to forward: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

// ─── IMAP Tools ─────────────────────────────────────────────────────────────

server.registerTool(
  "list_folders",
  {
    description: "List available email folders/mailboxes with message counts",
    inputSchema: {},
  },
  async () => {
    debugLog("[Tool] Executing tool: list_folders");

    try {
      const folders = await imapService.listFolders();
      const formatted = folders
        .map((f) => {
          const use = f.specialUse ? ` (${f.specialUse})` : "";
          return `${f.path}${use} — ${f.messages} messages, ${f.unseen} unread`;
        })
        .join("\n");

      return {
        content: [{ type: "text" as const, text: formatted || "No folders found." }],
      };
    } catch (error) {
      console.error(`[Error] Failed to list folders: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to list folders: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "list_messages",
  {
    description:
      "List recent messages from an email folder. Returns subject, sender, date, and flags for each message.",
    inputSchema: {
      folder: z.string().optional().default("INBOX").describe("Folder path to list messages from (default: INBOX)"),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .default(20)
        .describe("Maximum number of messages to return (default: 20, max: 100)"),
      beforeUid: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe(
          "Fetch messages with UIDs before this value (for pagination). Pass the smallest UID from the previous page.",
        ),
    },
  },
  async ({ folder, limit, beforeUid }) => {
    debugLog(
      `[Tool] Executing tool: list_messages (folder=${folder}, limit=${limit}${beforeUid ? `, beforeUid=${beforeUid}` : ""})`,
    );

    try {
      const messages = await imapService.listMessages(folder, limit, beforeUid);
      if (messages.length === 0) {
        return { content: [{ type: "text" as const, text: `No messages in ${folder}.` }] };
      }

      const formatted = messages
        .map((m) => {
          const flags = m.flags.length > 0 ? ` [${m.flags.join(", ")}]` : "";
          return `UID ${m.uid} | ${m.date} | From: ${m.from} | ${m.subject}${flags}`;
        })
        .join("\n");

      return {
        content: [{ type: "text" as const, text: `${messages.length} messages in ${folder}:\n\n${formatted}` }],
      };
    } catch (error) {
      console.error(`[Error] Failed to list messages: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to list messages: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "read_message",
  {
    description: "Read a specific email message by UID. Returns full headers and body content.",
    inputSchema: {
      uid: z.number().int().min(1).describe("Message UID (use list_messages or search_messages to find UIDs)"),
      folder: z.string().optional().default("INBOX").describe("Folder path containing the message (default: INBOX)"),
    },
  },
  async ({ uid, folder }) => {
    debugLog(`[Tool] Executing tool: read_message (uid=${uid}, folder=${folder})`);

    try {
      const msg = await imapService.readMessage(folder, uid);
      const parts: string[] = [
        `Subject: ${msg.subject}`,
        `From: ${msg.from}`,
        `To: ${msg.to}`,
        msg.cc ? `CC: ${msg.cc}` : "",
        `Date: ${msg.date}`,
        `Message-ID: ${msg.messageId}`,
        `Flags: ${msg.flags.join(", ") || "none"}`,
      ];

      if (msg.attachments.length > 0) {
        parts.push("");
        parts.push(`--- Attachments (${msg.attachments.length}) ---`);
        for (const att of msg.attachments) {
          parts.push(`  [${att.partNumber}] ${att.filename} (${att.contentType}, ${att.size} bytes)`);
        }
      }

      parts.push("");
      parts.push("--- Body ---");
      parts.push(msg.text || msg.html || "(no content)");

      return {
        content: [{ type: "text" as const, text: parts.join("\n") }],
      };
    } catch (error) {
      console.error(`[Error] Failed to read message: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to read message: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "download_attachment",
  {
    description:
      "Download an email attachment by part number. Use read_message first to see available attachments and their part numbers. Returns base64-encoded content.",
    inputSchema: {
      uid: z.number().int().min(1).describe("Message UID"),
      folder: z.string().optional().default("INBOX").describe("Folder containing the message (default: INBOX)"),
      partNumber: z
        .string()
        .min(1)
        .regex(/^\d+(\.\d+)*$/, "Invalid MIME part number format")
        .describe("MIME part number of the attachment (from read_message output)"),
    },
  },
  async ({ uid, folder, partNumber }) => {
    debugLog(`[Tool] Executing tool: download_attachment (uid=${uid}, folder=${folder}, part=${partNumber})`);

    try {
      const attachment = await imapService.downloadAttachment(folder, uid, partNumber);

      return {
        content: [
          {
            type: "text" as const,
            text: `Attachment: ${attachment.filename} (${attachment.contentType})\nContent (base64):\n${attachment.content}`,
          },
        ],
      };
    } catch (error) {
      console.error(`[Error] Failed to download attachment: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to download attachment: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "search_messages",
  {
    description:
      "Search for messages in a folder by various criteria (sender, subject, date, flags). Returns matching message summaries.",
    inputSchema: {
      folder: z.string().optional().default("INBOX").describe("Folder to search in (default: INBOX)"),
      from: z.string().optional().describe("Filter by sender email address or name"),
      to: z.string().optional().describe("Filter by recipient email address"),
      subject: z.string().optional().describe("Filter by subject (substring match)"),
      body: z.string().optional().describe("Filter by body content (substring match)"),
      since: z.string().optional().describe("Messages since this date (YYYY-MM-DD)"),
      before: z.string().optional().describe("Messages before this date (YYYY-MM-DD)"),
      seen: z.boolean().optional().describe("Filter by read status: true=read, false=unread"),
      flagged: z.boolean().optional().describe("Filter by flagged/starred status"),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .default(20)
        .describe("Maximum results to return (default: 20, max: 100)"),
    },
  },
  async ({ folder, from, to, subject, body, since, before, seen, flagged, limit }) => {
    debugLog(`[Tool] Executing tool: search_messages (folder=${folder})`);

    try {
      const messages = await imapService.searchMessages(
        folder,
        { from, to, subject, body, since, before, seen, flagged },
        limit,
      );

      if (messages.length === 0) {
        return { content: [{ type: "text" as const, text: "No messages found matching your criteria." }] };
      }

      const formatted = messages
        .map((m) => {
          const flags = m.flags.length > 0 ? ` [${m.flags.join(", ")}]` : "";
          return `UID ${m.uid} | ${m.date} | From: ${m.from} | ${m.subject}${flags}`;
        })
        .join("\n");

      return {
        content: [{ type: "text" as const, text: `${messages.length} messages found:\n\n${formatted}` }],
      };
    } catch (error) {
      console.error(`[Error] Failed to search messages: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to search messages: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

// ─── IMAP Mailbox Management Tools ──────────────────────────────────────────

server.registerTool(
  "move_message",
  {
    description: "Move an email message to a different folder",
    inputSchema: {
      uid: z.number().int().min(1).describe("Message UID (use list_messages or search_messages to find UIDs)"),
      folder: z.string().optional().default("INBOX").describe("Source folder (default: INBOX)"),
      destination: z.string().min(1).describe("Destination folder path (e.g. Archive, Trash, Spam)"),
    },
  },
  async ({ uid, folder, destination }) => {
    debugLog(`[Tool] Executing tool: move_message (uid=${uid}, ${folder} → ${destination})`);

    try {
      const success = await imapService.moveMessage(folder, uid, destination);
      if (!success) {
        return {
          content: [
            { type: "text" as const, text: `Failed to move message UID ${uid} — the server returned no confirmation.` },
          ],
          isError: true,
        };
      }
      return {
        content: [{ type: "text" as const, text: `Message UID ${uid} moved from ${folder} to ${destination}.` }],
      };
    } catch (error) {
      console.error(`[Error] Failed to move message: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to move message: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "delete_message",
  {
    description: "Permanently delete an email message",
    inputSchema: {
      uid: z.number().int().min(1).describe("Message UID (use list_messages or search_messages to find UIDs)"),
      folder: z.string().optional().default("INBOX").describe("Folder containing the message (default: INBOX)"),
    },
  },
  async ({ uid, folder }) => {
    debugLog(`[Tool] Executing tool: delete_message (uid=${uid}, folder=${folder})`);

    try {
      const success = await imapService.deleteMessage(folder, uid);
      if (!success) {
        return {
          content: [
            {
              type: "text" as const,
              text: `Failed to delete message UID ${uid} — the server returned no confirmation.`,
            },
          ],
          isError: true,
        };
      }
      return {
        content: [{ type: "text" as const, text: `Message UID ${uid} permanently deleted from ${folder}.` }],
      };
    } catch (error) {
      console.error(`[Error] Failed to delete message: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to delete message: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

server.registerTool(
  "update_message_flags",
  {
    description:
      "Add or remove flags on an email message. Common flags: \\\\Seen (read), \\\\Flagged (starred), \\\\Answered, \\\\Draft, \\\\Deleted",
    inputSchema: {
      uid: z.number().int().min(1).describe("Message UID"),
      folder: z.string().optional().default("INBOX").describe("Folder containing the message (default: INBOX)"),
      flagsToAdd: z
        .array(
          z.string().refine((f) => {
            validateImapFlag(f);
            return true;
          }, "Invalid IMAP flag format"),
        )
        .optional()
        .default([])
        .describe('Flags to add (e.g. ["\\\\Seen", "\\\\Flagged"])'),
      flagsToRemove: z
        .array(
          z.string().refine((f) => {
            validateImapFlag(f);
            return true;
          }, "Invalid IMAP flag format"),
        )
        .optional()
        .default([])
        .describe('Flags to remove (e.g. ["\\\\Seen"])'),
    },
  },
  async ({ uid, folder, flagsToAdd, flagsToRemove }) => {
    debugLog(`[Tool] Executing tool: update_message_flags (uid=${uid}, folder=${folder})`);

    if (flagsToAdd.length === 0 && flagsToRemove.length === 0) {
      return {
        content: [{ type: "text" as const, text: "No flags specified to add or remove." }],
        isError: true,
      };
    }

    try {
      await imapService.updateFlags(folder, uid, flagsToAdd, flagsToRemove);

      const parts: string[] = [];
      if (flagsToAdd.length > 0) parts.push(`added: ${flagsToAdd.join(", ")}`);
      if (flagsToRemove.length > 0) parts.push(`removed: ${flagsToRemove.join(", ")}`);

      return {
        content: [{ type: "text" as const, text: `Flags updated on UID ${uid} in ${folder}: ${parts.join("; ")}.` }],
      };
    } catch (error) {
      console.error(`[Error] Failed to update flags: ${error instanceof Error ? error.message : String(error)}`);
      return {
        content: [{ type: "text" as const, text: `Failed to update flags: ${sanitizeError(error)}` }],
        isError: true,
      };
    }
  },
);

// ─── Server Startup ─────────────────────────────────────────────────────────

async function main() {
  debugLog("[Setup] Starting Proton Mail MCP server...");

  // Warn if connecting to remote hosts without TLS
  const localhosts = ["127.0.0.1", "localhost", "::1"];
  if (!localhosts.includes(IMAP_HOST) && !IMAP_SECURE) {
    console.error(
      "[Warning] IMAP connection to remote host without TLS. Set IMAP_SECURE=true for encrypted connections.",
    );
  }
  if (!localhosts.includes(PROTONMAIL_HOST) && !PROTONMAIL_SECURE && PROTONMAIL_PORT !== 587) {
    console.error(
      "[Warning] SMTP connection to remote host without TLS. Set PROTONMAIL_SECURE=true for encrypted connections.",
    );
  }

  // Non-fatal SMTP verification — warn but continue so the server stays up
  try {
    await emailService.verifyConnection();
  } catch (error) {
    console.error(`[Warning] SMTP verification failed, will retry on first send: ${sanitizeError(error)}`);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);

  debugLog("[Setup] Proton Mail MCP server started successfully");
}

// Set up error handling
process.on("uncaughtException", (error) => {
  console.error(`[Error] Uncaught exception: ${error.message}`);
  process.exit(1);
});

process.on("unhandledRejection", (reason) => {
  console.error(`[Error] Unhandled rejection: ${reason instanceof Error ? reason.message : String(reason)}`);
  process.exit(1);
});

// Start the server
main().catch((error) => {
  console.error(`[Error] Server error: ${error instanceof Error ? error.message : String(error)}`);
  process.exit(1);
});
