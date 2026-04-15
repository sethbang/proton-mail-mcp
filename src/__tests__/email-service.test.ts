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
        attachments: [
          { filename: "test.txt", content: base64Content, contentType: "text/plain" },
        ],
      });

      const call = mock.sendMail.mock.calls[0][0];
      expect(call.attachments).toHaveLength(1);
      expect(call.attachments[0].filename).toBe("test.txt");
      expect(call.attachments[0].contentType).toBe("text/plain");
      expect(Buffer.isBuffer(call.attachments[0].content)).toBe(true);
      expect(call.attachments[0].content.toString()).toBe("hello world");
    });

    it("re-throws errors from sendMail", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();
      mock.sendMail.mockRejectedValueOnce(new Error("SMTP connection refused"));

      await expect(
        service.sendEmail({ to: "x@example.com", subject: "Test", body: "Test" }),
      ).rejects.toThrow("SMTP connection refused");
    });
  });

  describe("verifyConnection", () => {
    it("resolves without error on success", async () => {
      const service = new EmailService(baseConfig);
      await expect(service.verifyConnection()).resolves.toBeUndefined();
    });

    it("re-throws on verification failure", async () => {
      const service = new EmailService(baseConfig);
      const mock = getTransporterMock();
      mock.verify.mockRejectedValueOnce(new Error("Connection refused"));

      await expect(service.verifyConnection()).rejects.toThrow("Connection refused");
    });
  });
});
