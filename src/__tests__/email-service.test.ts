import { describe, it, expect, vi, beforeEach } from "vitest";
import nodemailer from "nodemailer";
import { EmailService } from "../email-service.js";
import type { EmailConfig } from "../email-service.js";

// Mock nodemailer
vi.mock("nodemailer", () => {
  const mockSendMail = vi.fn().mockResolvedValue({ messageId: "<test-id@protonmail.ch>" });
  const mockVerify = vi.fn().mockResolvedValue(true);
  return {
    default: {
      createTransport: vi.fn(() => ({
        sendMail: mockSendMail,
        verify: mockVerify,
      })),
    },
  };
});

function getTransporterMock() {
  const calls = vi.mocked(nodemailer.createTransport).mock.results;
  return calls[calls.length - 1].value as {
    sendMail: ReturnType<typeof vi.fn>;
    verify: ReturnType<typeof vi.fn>;
  };
}

const baseConfig: EmailConfig = {
  host: "smtp.protonmail.ch",
  port: 587,
  secure: false,
  auth: { user: "test@protonmail.com", pass: "test-password" },
};

describe("EmailService", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("constructor", () => {
    it("creates transporter with correct config", () => {
      new EmailService(baseConfig);
      expect(nodemailer.createTransport).toHaveBeenCalledWith(
        expect.objectContaining({
          host: "smtp.protonmail.ch",
          port: 587,
          secure: false,
          auth: { user: "test@protonmail.com", pass: "test-password" },
        }),
      );
    });

    it("logs initialization when debug is true", () => {
      const spy = vi.spyOn(console, "error").mockImplementation(() => {});
      new EmailService({ ...baseConfig, debug: true });
      expect(spy).toHaveBeenCalledWith("[Setup] Initializing email service...");
      spy.mockRestore();
    });

    it("applies timeout defaults", () => {
      new EmailService(baseConfig);
      expect(nodemailer.createTransport).toHaveBeenCalledWith(
        expect.objectContaining({
          connectionTimeout: 30_000,
          greetingTimeout: 30_000,
          socketTimeout: 60_000,
        }),
      );
    });

    it("applies custom timeouts", () => {
      new EmailService({ ...baseConfig, connectionTimeout: 5_000, socketTimeout: 10_000 });
      expect(nodemailer.createTransport).toHaveBeenCalledWith(
        expect.objectContaining({
          connectionTimeout: 5_000,
          socketTimeout: 10_000,
        }),
      );
    });
  });

  describe("sendEmail", () => {
    it("sends plain text email with text field set and html undefined", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Test Subject",
        body: "Hello, World!",
      });

      expect(mock.sendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          from: "test@protonmail.com",
          to: "recipient@example.com",
          subject: "Test Subject",
          text: "Hello, World!",
          html: undefined,
        }),
      );
    });

    it("sends HTML email with both html and plaintext fallback", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "HTML Test",
        body: "<p>Hello</p><p>World</p>",
        isHtml: true,
      });

      const call = mock.sendMail.mock.calls[0][0];
      expect(call.html).toBe("<p>Hello</p><p>World</p>");
      expect(call.text).toBe("Hello\n\nWorld");
    });

    it("sanitizes fromName to prevent header injection", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Test",
        body: "Hello",
        fromName: 'John "Bobby" Doe',
      });

      expect(mock.sendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          from: '"John Bobby Doe" <test@protonmail.com>',
        }),
      );
    });

    it("passes cc and bcc through", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "CC Test",
        body: "Hello",
        cc: "cc@example.com",
        bcc: "bcc@example.com",
      });

      expect(mock.sendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          cc: "cc@example.com",
          bcc: "bcc@example.com",
        }),
      );
    });

    it("returns nodemailer send info", async () => {
      const service = new EmailService(baseConfig);
      const result = await service.sendEmail({
        to: "recipient@example.com",
        subject: "Test",
        body: "Test",
      });

      expect(result).toEqual({ messageId: "<test-id@protonmail.ch>" });
    });

    it("passes inReplyTo and references headers", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Re: Test",
        body: "Reply body",
        inReplyTo: "<original-id@example.com>",
        references: "<original-id@example.com>",
      });

      expect(mock.sendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          inReplyTo: "<original-id@example.com>",
          references: "<original-id@example.com>",
        }),
      );
    });

    it("converts base64 attachments to Buffers", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();
      const base64Content = Buffer.from("hello world").toString("base64");

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Test",
        body: "See attached",
        attachments: [{ filename: "test.txt", content: base64Content, contentType: "text/plain" }],
      });

      const call = mock.sendMail.mock.calls[0][0];
      expect(call.attachments).toHaveLength(1);
      expect(call.attachments[0].filename).toBe("test.txt");
      expect(call.attachments[0].contentType).toBe("text/plain");
      expect(Buffer.isBuffer(call.attachments[0].content)).toBe(true);
      expect(call.attachments[0].content.toString()).toBe("hello world");
    });

    it("logs send and success when debug is true", async () => {
      const spy = vi.spyOn(console, "error").mockImplementation(() => {});
      const service = new EmailService({ ...baseConfig, debug: true });

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Debug Test",
        body: "Hello",
      });

      expect(spy).toHaveBeenCalledWith("[Email] Sending email to: recipient@example.com");
      expect(spy).toHaveBeenCalledWith(expect.stringContaining("[Email] Email sent successfully:"));
      spy.mockRestore();
    });

    it("strips HTML entities in plaintext fallback", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Entity Test",
        body: "<p>&nbsp;&amp;&lt;&gt;&quot;</p><br/><p>line2</p>",
        isHtml: true,
      });

      const call = mock.sendMail.mock.calls[0][0];
      expect(call.text).toBe('&<>"\n\nline2');
    });

    it("strips CRLF and angle brackets from fromName to prevent header injection", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Test",
        body: "Hello",
        fromName: "Evil\r\nBcc: attacker@evil.com",
      });

      expect(mock.sendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          from: '"EvilBcc: attacker@evil.com" <test@protonmail.com>',
        }),
      );
    });

    it("sanitizes path traversal from attachment filenames", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();
      const base64Content = Buffer.from("data").toString("base64");

      await service.sendEmail({
        to: "recipient@example.com",
        subject: "Test",
        body: "See attached",
        attachments: [
          { filename: "../../etc/passwd", content: base64Content, contentType: "text/plain" },
          { filename: "normal.txt", content: base64Content, contentType: "text/plain" },
        ],
      });

      const call = mock.sendMail.mock.calls[0][0];
      expect(call.attachments[0].filename).toBe("passwd");
      expect(call.attachments[1].filename).toBe("normal.txt");
    });

    it("re-throws errors from sendMail", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();
      mock.sendMail.mockRejectedValueOnce(new Error("SMTP connection refused"));

      await expect(service.sendEmail({ to: "x@example.com", subject: "Test", body: "Test" })).rejects.toThrow(
        "SMTP connection refused",
      );
    });
  });

  describe("verifyConnection", () => {
    it("resolves without error on success", async () => {
      const service = new EmailService(baseConfig);
      await expect(service.verifyConnection()).resolves.toBeUndefined();
    });

    it("logs verification when debug is true", async () => {
      const spy = vi.spyOn(console, "error").mockImplementation(() => {});
      const service = new EmailService({ ...baseConfig, debug: true });
      await service.verifyConnection();
      expect(spy).toHaveBeenCalledWith("[Setup] Verifying SMTP connection...");
      expect(spy).toHaveBeenCalledWith("[Setup] SMTP connection verified successfully");
      spy.mockRestore();
    });

    it("re-throws on verification failure", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();
      mock.verify.mockRejectedValueOnce(new Error("Connection refused"));

      await expect(service.verifyConnection()).rejects.toThrow("Connection refused");
    });
  });
});
