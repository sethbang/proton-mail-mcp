/**
 * End-to-end smoke test for the MCP server surface.
 *
 * Every other test file exercises a service module in isolation. This one
 * spawns the built server as a real child process, speaks MCP over stdio, and
 * asserts on what a client actually sees. It exists to guard the SDK v2
 * migration: `npm run build && npm test` passing tells you nothing about
 * whether 31 tool registrations survived a schema-library change, and the
 * failure mode that matters most — losing `.describe()` field descriptions
 * during Zod-to-JSON-Schema conversion — degrades every tool call without
 * breaking a single type.
 *
 * No network: `tools/list` never reaches SMTP or IMAP, and both hosts point at
 * a closed local port so the non-fatal startup verification fails fast
 * (ECONNREFUSED) instead of burning the connection timeout.
 */

import { describe, it, expect, beforeAll } from "vitest";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { Client } from "@modelcontextprotocol/client";
import type { Tool } from "@modelcontextprotocol/client";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const PROJECT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const SERVER_ENTRY = path.join(PROJECT_ROOT, "build", "index.js");

/** Spawn timeout per connection — generous for cold CI runners (local: ~300ms). */
const SPAWN_TIMEOUT_MS = 20_000;

/** Tools registered with no env gating at all. */
const DEFAULT_TOOLS = [
  "send_email",
  "reply_email",
  "reply_all_email",
  "forward_email",
  "list_folders",
  "list_messages",
  "read_message",
  "list_attachments",
  "download_attachment",
  "search_messages",
  "move_message",
  "delete_message",
  "update_message_flags",
  "get_thread",
  "mark_all_read",
  "save_draft",
  "bulk_move",
  "bulk_delete",
  "bulk_update_flags",
  "create_folder",
  "create_label",
  "rename_folder",
  "delete_folder",
  "update_message_labels",
  "bulk_update_labels",
  "count_messages",
  "folder_stats",
  "top_senders",
  "move_thread",
  "delete_thread",
  "flag_thread",
] as const;

/** The subset that survives READONLY=true — nothing here mutates a mailbox. */
const READONLY_MODE_TOOLS = [
  "list_folders",
  "list_messages",
  "read_message",
  "list_attachments",
  "download_attachment",
  "search_messages",
  "get_thread",
  "count_messages",
  "folder_stats",
  "top_senders",
] as const;

/**
 * Tools that advertise `readOnlyHint: true` — no side effects *anywhere*, not
 * just in the mailbox.
 *
 * Deliberately not the same list as READONLY_MODE_TOOLS. READONLY gates mailbox
 * mutation; `readOnlyHint` is a stronger claim to the client.
 * `download_attachment` passes the first bar and fails the second: it never
 * touches the mailbox, but it can write bytes to disk via `saveTo` — see commit
 * 35b3035. Any future tool that is mailbox-safe but touches the filesystem,
 * network, or clock belongs in the first list and not this one.
 */
const SIDE_EFFECT_FREE_TOOLS = [
  "list_folders",
  "list_messages",
  "read_message",
  "list_attachments",
  "search_messages",
  "get_thread",
  "count_messages",
  "folder_stats",
  "top_senders",
] as const;

/**
 * Every nested (non-top-level) parameter that carries `.describe()` text today.
 *
 * The list is short because `searchCriteriaSchema` documents only two of its
 * fourteen fields — a pre-existing authoring gap (60 nested params have no
 * description), not conversion loss. Pinning the exact set means the migration
 * canary covers nested schemas too: JSON Schema conversion can drop
 * descriptions inside `items` / nested `properties` while leaving top-level
 * ones intact, and a top-level-only assertion would sail right past that.
 */
const DESCRIBED_NESTED_PARAMS = [
  "bulk_delete.match.attachmentName",
  "bulk_delete.match.attachmentType",
  "bulk_move.match.attachmentName",
  "bulk_move.match.attachmentType",
  "bulk_update_flags.match.attachmentName",
  "bulk_update_flags.match.attachmentType",
  "bulk_update_labels.match.attachmentName",
  "bulk_update_labels.match.attachmentType",
  "count_messages.match.attachmentName",
  "count_messages.match.attachmentType",
  "send_email.attachments[].content",
  "send_email.attachments[].contentType",
  "send_email.attachments[].filename",
] as const;

type JsonSchemaNode = {
  description?: string;
  properties?: Record<string, JsonSchemaNode>;
  items?: JsonSchemaNode;
};

/** Collect dotted paths of every described parameter below the top level. */
function describedNestedParams(tools: Tool[]): string[] {
  const found: string[] = [];
  const walk = (node: JsonSchemaNode | undefined, pathStr: string, depth: number): void => {
    if (!node || typeof node !== "object") return;
    for (const [name, child] of Object.entries(node.properties ?? {})) {
      const childPath = `${pathStr}.${name}`;
      if (depth > 0 && child?.description) found.push(childPath);
      walk(child, childPath, depth + 1);
      if (child?.items) walk(child.items, `${childPath}[]`, depth + 1);
    }
  };
  for (const tool of tools) walk(tool.inputSchema as JsonSchemaNode, tool.name, 0);
  return found.sort();
}

type ClientOptions = ConstructorParameters<typeof Client>[1];

/**
 * Connect a real MCP client to a freshly spawned server and report both the
 * era the connection negotiated and the tools it sees.
 *
 * The credentials are fake by design; nothing in `tools/list` authenticates.
 */
async function connectAndList(
  clientOptions?: ClientOptions,
  extraEnv: Record<string, string> = {},
): Promise<{ era: string | undefined; tools: Tool[] }> {
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [SERVER_ENTRY],
    env: {
      PATH: process.env.PATH ?? "",
      PROTONMAIL_USERNAME: "smoke@example.com",
      PROTONMAIL_PASSWORD: "smoke-password",
      // Port 1 on loopback refuses immediately, so the non-fatal SMTP
      // verifyConnection() in main() fails in milliseconds rather than
      // waiting out connectionTimeout (30s).
      PROTONMAIL_HOST: "127.0.0.1",
      PROTONMAIL_PORT: "1",
      IMAP_HOST: "127.0.0.1",
      IMAP_PORT: "1",
      ...extraEnv,
    },
    stderr: "pipe",
  });

  const client = new Client({ name: "server-smoke-test", version: "0.0.0" }, clientOptions);
  try {
    await client.connect(transport);
    const { tools } = await client.listTools();
    return { era: client.getProtocolEra(), tools };
  } finally {
    await client.close();
  }
}

/** Tools as seen by a default client, which negotiates the 2025 era. */
async function listTools(extraEnv: Record<string, string> = {}): Promise<Tool[]> {
  return (await connectAndList(undefined, extraEnv)).tools;
}

describe("MCP server surface (stdio, end-to-end)", () => {
  beforeAll(() => {
    if (!fs.existsSync(SERVER_ENTRY)) {
      throw new Error(`Missing ${SERVER_ENTRY}. Run \`npm run build\` before \`npm test\`.`);
    }
  });

  describe("default registration", () => {
    let tools: Tool[];

    beforeAll(async () => {
      tools = await listTools();
    }, SPAWN_TIMEOUT_MS);

    it("registers exactly the expected tools", () => {
      expect(tools.map((t) => t.name).sort()).toEqual([...DEFAULT_TOOLS].sort());
    });

    it("does not register empty_folder unless opted in", () => {
      expect(tools.some((t) => t.name === "empty_folder")).toBe(false);
    });

    it("gives every tool a description and an object input schema", () => {
      for (const tool of tools) {
        expect(tool.description, `${tool.name} has no description`).toBeTruthy();
        expect(tool.inputSchema?.type, `${tool.name} input schema is not an object`).toBe("object");
      }
    });

    /**
     * The migration canary. A Zod below 4.2.0 (or a schema authored by a
     * different Zod instance than the one converting it) silently drops
     * `.describe()` text during JSON Schema conversion — the build passes, the
     * server starts, and every tool call gets worse. Asserting the invariant
     * directly is what makes that visible.
     */
    it("preserves .describe() text on every top-level parameter", () => {
      const undescribed: string[] = [];
      for (const tool of tools) {
        const properties = (tool.inputSchema?.properties ?? {}) as Record<string, { description?: string }>;
        for (const [name, schema] of Object.entries(properties)) {
          if (!schema?.description) undescribed.push(`${tool.name}.${name}`);
        }
      }
      expect(undescribed).toEqual([]);
    });

    /**
     * Same canary, one level down. Conversion can preserve top-level
     * descriptions while dropping the ones inside `items` and nested
     * `properties`, so this pins the exact set rather than a count.
     */
    it("preserves .describe() text on nested parameters", () => {
      expect(describedNestedParams(tools)).toEqual([...DESCRIBED_NESTED_PARAMS].sort());
    });

    it("preserves annotations and required fields on send_email", () => {
      const send = tools.find((t) => t.name === "send_email");
      expect(send?.annotations).toMatchObject({
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
      });
      expect(send?.inputSchema?.required).toEqual(["to", "subject"]);
      expect(send?.inputSchema?.properties?.to).toMatchObject({
        description: expect.stringContaining("Recipient email address"),
      });
    });

    it("marks side-effect-free tools with readOnlyHint", () => {
      for (const name of SIDE_EFFECT_FREE_TOOLS) {
        const tool = tools.find((t) => t.name === name);
        expect(tool?.annotations?.readOnlyHint, `${name} is not annotated readOnlyHint`).toBe(true);
      }
    });

    it("withholds readOnlyHint from download_attachment, which can write to disk", () => {
      const tool = tools.find((t) => t.name === "download_attachment");
      expect(tool?.annotations?.readOnlyHint).toBe(false);
    });
  });

  describe("READONLY=true", () => {
    it(
      "registers only non-mutating tools",
      async () => {
        const tools = await listTools({ READONLY: "true" });
        expect(tools.map((t) => t.name).sort()).toEqual([...READONLY_MODE_TOOLS].sort());
      },
      SPAWN_TIMEOUT_MS,
    );
  });

  describe("ALLOW_EMPTY_FOLDER=true", () => {
    it(
      "adds empty_folder to the default set",
      async () => {
        const tools = await listTools({ ALLOW_EMPTY_FOLDER: "true" });
        expect(tools.map((t) => t.name).sort()).toEqual([...DEFAULT_TOOLS, "empty_folder"].sort());
      },
      SPAWN_TIMEOUT_MS,
    );

    it(
      "stays disabled under READONLY",
      async () => {
        const tools = await listTools({ ALLOW_EMPTY_FOLDER: "true", READONLY: "true" });
        expect(tools.some((t) => t.name === "empty_folder")).toBe(false);
      },
      SPAWN_TIMEOUT_MS,
    );
  });

  /**
   * `serveStdio` decides the protocol era once per connection, and its default
   * (`legacy: 'serve'`) answers both eras from the same factory. These are the
   * assertions that would catch someone passing `{ legacy: 'reject' }`, which
   * would silently cut off every client that has not adopted 2026-07-28.
   *
   * The tool surface is asserted on both paths deliberately: a server can
   * negotiate an era correctly and still serve a different set of tools on it.
   */
  describe("protocol era negotiation", () => {
    it(
      "serves the 2025 era to a default client",
      async () => {
        const { era, tools } = await connectAndList();
        expect(era).toBe("legacy");
        expect(tools).toHaveLength(DEFAULT_TOOLS.length);
      },
      SPAWN_TIMEOUT_MS,
    );

    it(
      "serves the 2026-07-28 era to a client pinned to it",
      async () => {
        const { era, tools } = await connectAndList({ versionNegotiation: { mode: { pin: "2026-07-28" } } });
        expect(era).toBe("modern");
        expect(tools).toHaveLength(DEFAULT_TOOLS.length);
      },
      SPAWN_TIMEOUT_MS,
    );

    it(
      "negotiates the modern era when a client probes with mode auto",
      async () => {
        const { era } = await connectAndList({ versionNegotiation: { mode: "auto" } });
        expect(era).toBe("modern");
      },
      SPAWN_TIMEOUT_MS,
    );

    it(
      "exposes an identical tool surface on both eras",
      async () => {
        const legacy = await connectAndList();
        const modern = await connectAndList({ versionNegotiation: { mode: { pin: "2026-07-28" } } });
        expect(legacy.era).toBe("legacy");
        expect(modern.era).toBe("modern");
        expect(modern.tools.map((t) => t.name).sort()).toEqual(legacy.tools.map((t) => t.name).sort());
      },
      SPAWN_TIMEOUT_MS,
    );
  });
});
