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
  console.error("[Error] Missing required environment variables: PROTONMAIL_USERNAME and PROTONMAIL_PASSWORD must be set");
  process.exit(1);
}

// Helper function for debug logging
function debugLog(message: string): void {
  if (DEBUG) {
    console.error(message);
  }
}

/**
 * Strip credential-like substrings from error messages before returning to MCP clients.
 * Full error details are still logged to stderr for debugging.
 */
function sanitizeError(error: unknown): string {
  const msg = error instanceof Error ? error.message : String(error);
  return msg
    .replace(/(?:user|pass|password|auth|credentials?)[=:\s]+"?[^\s"]+/gi, "[REDACTED]")
    .replace(/(?:smtp|imap):\/\/[^\s]+/gi, "[REDACTED_URL]");
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
  version: "0.2.0",
});

// Basic email format check for comma-separated address fields
const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
function validateAddresses(value: string): boolean {
  return value.split(",").every((addr) => emailRegex.test(addr.trim()));
}

// ─── SMTP Tools ─────────────────────────────────────────────────────────────

server.registerTool(
  "send_email",
  {
    description: "Send an email using Proton Mail SMTP",
    inputSchema: {
      to: z.string()
        .min(1, "Recipient is required")
        .max(10_000, "To field too long")
        .refine(validateAddresses, "Each 'to' address must be a valid email")
        .describe("Recipient email address(es). Multiple addresses can be separated by commas."),
      subject: z.string()
        .min(1, "Subject is required")
        .max(998, "Subject exceeds RFC 5322 line length limit")
        .describe("Email subject line"),
      body: z.string()
        .min(1, "Body is required")
        .max(500_000, "Body too large")
        .describe("Email body content (can be plain text or HTML)"),
      isHtml: z.boolean().optional().default(false)
        .describe("Whether the body contains HTML content"),
      cc: z.string()
        .max(10_000, "CC field too long")
        .refine(validateAddresses, "Each CC address must be a valid email")
        .optional()
        .describe("CC recipient(s), separated by commas"),
      bcc: z.string()
        .max(10_000, "BCC field too long")
        .refine(validateAddresses, "Each BCC address must be a valid email")
        .optional()
        .describe("BCC recipient(s), separated by commas"),
      replyTo: z.string()
        .max(10_000, "Reply-To field too long")
        .refine(validateAddresses, "Reply-To must be a valid email")
        .optional()
        .describe("Reply-To email address"),
      fromName: z.string()
        .max(200, "From name too long")
        .optional()
        .describe("Display name for the From field"),
    },
  },
  async ({ to, subject, body, isHtml, cc, bcc, replyTo, fromName }) => {
    debugLog(`[Tool] Executing tool: send_email`);

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
      });

      return {
        content: [{
          type: "text" as const,
          text: `Email sent successfully to ${to}${cc ? ` with CC to ${cc}` : ""}${bcc ? ` and BCC to ${bcc}` : ""}.`,
        }],
      };
    } catch (error) {
      console.error(`[Error] Failed to send email: ${error instanceof Error ? error.message : String(error)}`);

      return {
        content: [{
          type: "text" as const,
          text: `Failed to send email: ${sanitizeError(error)}`,
        }],
        isError: true,
      };
    }
  }
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
  }
);

server.registerTool(
  "list_messages",
  {
    description: "List recent messages from an email folder. Returns subject, sender, date, and flags for each message.",
    inputSchema: {
      folder: z.string().optional().default("INBOX")
        .describe("Folder path to list messages from (default: INBOX)"),
      limit: z.number().int().min(1).max(100).optional().default(20)
        .describe("Maximum number of messages to return (default: 20, max: 100)"),
    },
  },
  async ({ folder, limit }) => {
    debugLog(`[Tool] Executing tool: list_messages (folder=${folder}, limit=${limit})`);

    try {
      const messages = await imapService.listMessages(folder, limit);
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
  }
);

server.registerTool(
  "read_message",
  {
    description: "Read a specific email message by UID. Returns full headers and body content.",
    inputSchema: {
      uid: z.number().int().min(1)
        .describe("Message UID (use list_messages or search_messages to find UIDs)"),
      folder: z.string().optional().default("INBOX")
        .describe("Folder path containing the message (default: INBOX)"),
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
        "",
        "--- Body ---",
        msg.text || msg.html || "(no content)",
      ].filter(Boolean);

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
  }
);

server.registerTool(
  "search_messages",
  {
    description: "Search for messages in a folder by various criteria (sender, subject, date, flags). Returns matching message summaries.",
    inputSchema: {
      folder: z.string().optional().default("INBOX")
        .describe("Folder to search in (default: INBOX)"),
      from: z.string().optional()
        .describe("Filter by sender email address or name"),
      to: z.string().optional()
        .describe("Filter by recipient email address"),
      subject: z.string().optional()
        .describe("Filter by subject (substring match)"),
      body: z.string().optional()
        .describe("Filter by body content (substring match)"),
      since: z.string().optional()
        .describe("Messages since this date (YYYY-MM-DD)"),
      before: z.string().optional()
        .describe("Messages before this date (YYYY-MM-DD)"),
      seen: z.boolean().optional()
        .describe("Filter by read status: true=read, false=unread"),
      flagged: z.boolean().optional()
        .describe("Filter by flagged/starred status"),
      limit: z.number().int().min(1).max(100).optional().default(20)
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
  }
);

// ─── Server Startup ─────────────────────────────────────────────────────────

async function main() {
  debugLog("[Setup] Starting Proton Mail MCP server...");

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
