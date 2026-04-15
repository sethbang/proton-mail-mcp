import nodemailer from "nodemailer";
import type SentMessageInfo from "nodemailer/lib/smtp-transport/index.js";

/**
 * Interface for email configuration
 */
export interface EmailConfig {
  host: string;
  port: number;
  secure: boolean;
  auth: {
    user: string;
    pass: string;
  };
  debug?: boolean;
  connectionTimeout?: number;
  greetingTimeout?: number;
  socketTimeout?: number;
}

/**
 * Interface for email message
 */
export interface EmailMessage {
  to: string;
  subject: string;
  body: string;
  isHtml?: boolean;
  cc?: string;
  bcc?: string;
  replyTo?: string;
  fromName?: string;
}

/**
 * Strip HTML tags and decode common entities to produce a plaintext fallback.
 */
function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * Email service for sending emails via SMTP
 */
export class EmailService {
  private transporter: nodemailer.Transporter;
  private fromEmail: string;
  private debug: boolean;

  /**
   * Create a new EmailService instance
   * @param config SMTP configuration
   */
  constructor(config: EmailConfig) {
    this.debug = config.debug || false;

    if (this.debug) {
      console.error("[Setup] Initializing email service...");
    }

    this.fromEmail = config.auth.user;
    this.transporter = nodemailer.createTransport({
      host: config.host,
      port: config.port,
      secure: config.secure,
      auth: {
        user: config.auth.user,
        pass: config.auth.pass,
      },
      connectionTimeout: config.connectionTimeout ?? 30_000,
      greetingTimeout: config.greetingTimeout ?? 30_000,
      socketTimeout: config.socketTimeout ?? 60_000,
    });
  }

  /**
   * Send an email
   * @param message Email message to send
   * @returns Promise resolving to the nodemailer send info
   */
  async sendEmail(message: EmailMessage): Promise<SentMessageInfo> {
    if (this.debug) {
      console.error(`[Email] Sending email to: ${message.to}`);
    }

    try {
      const safeName = message.fromName?.replace(/["\\]/g, "") || "";
      const from = safeName ? `"${safeName}" <${this.fromEmail}>` : this.fromEmail;
      const info = await this.transporter.sendMail({
        from,
        to: message.to,
        cc: message.cc,
        bcc: message.bcc,
        replyTo: message.replyTo,
        subject: message.subject,
        text: message.isHtml ? stripHtml(message.body) : message.body,
        html: message.isHtml ? message.body : undefined,
      });

      if (this.debug) {
        console.error(`[Email] Email sent successfully: ${info.messageId}`);
      }
      return info;
    } catch (error) {
      // Error is logged here for diagnostics; caller (index.ts) sanitizes before returning to MCP client
      console.error(`[Error] Failed to send email: ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /**
   * Verify SMTP connection
   */
  async verifyConnection(): Promise<void> {
    if (this.debug) {
      console.error("[Setup] Verifying SMTP connection...");
    }

    try {
      await this.transporter.verify();
      if (this.debug) {
        console.error("[Setup] SMTP connection verified successfully");
      }
    } catch (error) {
      console.error(`[Error] SMTP connection verification failed: ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }
}
