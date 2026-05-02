import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { registerReadTools } from "../src/tools/read.js";
import { registerDraftTools } from "../src/tools/drafts.js";
import { registerOrganizeTools } from "../src/tools/organize.js";
import type { MessagesApi } from "../src/graph/messages.js";
import type { DraftsApi } from "../src/graph/drafts.js";

let originalCacheDir: string | undefined;
beforeAll(() => {
  originalCacheDir = process.env.OUTLOOK_MCP_CACHE_DIR;
  process.env.OUTLOOK_MCP_CACHE_DIR = mkdtempSync(join(tmpdir(), "outlook-mcp-tools-"));
});
afterAll(() => {
  if (originalCacheDir === undefined) delete process.env.OUTLOOK_MCP_CACHE_DIR;
  else process.env.OUTLOOK_MCP_CACHE_DIR = originalCacheDir;
});

interface BuiltHarness {
  server: McpServer;
  client: Client;
  toolNames: string[];
}

async function buildHarness(opts: {
  readOnly: boolean;
  enableOrganize: boolean;
  messages?: Partial<MessagesApi>;
  drafts?: Partial<DraftsApi>;
}): Promise<BuiltHarness> {
  const server = new McpServer({ name: "test", version: "0.0.0" });
  const messagesStub = (opts.messages ?? {}) as MessagesApi;
  const draftsStub = (opts.drafts ?? {}) as DraftsApi;

  registerReadTools(server, messagesStub);
  if (!opts.readOnly) {
    registerDraftTools(server, draftsStub);
    if (opts.enableOrganize) registerOrganizeTools(server, messagesStub);
  }

  const client = new Client({ name: "test-client", version: "0.0.0" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(a), client.connect(b)]);

  const list = await client.listTools();
  const toolNames = list.tools.map((t) => t.name).sort();
  return { server, client, toolNames };
}

describe("tool wiring", () => {
  it("read-only mode exposes only read tools, no send/draft/organize", async () => {
    const { toolNames } = await buildHarness({ readOnly: true, enableOrganize: false });
    expect(toolNames).toEqual(["get_message", "list_folders", "list_messages", "search_messages"]);
    expect(toolNames.some((n) => n.includes("send"))).toBe(false);
    expect(toolNames.some((n) => n.includes("delete") && !n.includes("draft"))).toBe(false);
  });

  it("default mode exposes read + draft tools but no send tool ever", async () => {
    const { toolNames } = await buildHarness({ readOnly: false, enableOrganize: false });
    expect(toolNames).toContain("create_draft");
    expect(toolNames).toContain("update_draft");
    expect(toolNames).toContain("list_drafts");
    expect(toolNames).toContain("delete_draft");
    expect(toolNames.some((n) => n.startsWith("send"))).toBe(false);
    expect(toolNames).not.toContain("move_message");
  });

  it("--enable-organize adds move/mark/flag/folder tools", async () => {
    const { toolNames } = await buildHarness({ readOnly: false, enableOrganize: true });
    expect(toolNames).toContain("move_message");
    expect(toolNames).toContain("mark_read");
    expect(toolNames).toContain("mark_unread");
    expect(toolNames).toContain("flag");
    expect(toolNames).toContain("unflag");
    expect(toolNames).toContain("create_folder");
    expect(toolNames).toContain("rename_folder");
    expect(toolNames.some((n) => n.startsWith("send"))).toBe(false);
  });
});

describe("zod input validation", () => {
  it("rejects bad email addresses on create_draft", async () => {
    const { client } = await buildHarness({
      readOnly: false,
      enableOrganize: false,
      drafts: {
        async create() {
          throw new Error("should not be called: validation must reject first");
        },
      },
    });
    const result = await client.callTool({
      name: "create_draft",
      arguments: { to: ["not-an-email"], subject: "x", body: "y" },
    });
    expect(result.isError).toBe(true);
  });

  it("rejects out-of-range top on list_messages", async () => {
    const { client } = await buildHarness({
      readOnly: true,
      enableOrganize: false,
      messages: {
        async listMessages() {
          throw new Error("should not be called");
        },
      },
    });
    const result = await client.callTool({
      name: "list_messages",
      arguments: { top: 999 },
    });
    expect(result.isError).toBe(true);
  });
});

describe("organize safety rails", () => {
  it("move_message refuses 'permanentdelete' destination", async () => {
    const { client } = await buildHarness({
      readOnly: false,
      enableOrganize: true,
      messages: {
        async moveMessage() {
          throw new Error("should not be called: destination must be rejected by zod");
        },
      },
    });
    const result = await client.callTool({
      name: "move_message",
      arguments: { id: "abc", destinationFolder: "permanentdelete" },
    });
    expect(result.isError).toBe(true);
    const text = (result.content as any[]).map((c) => c.text).join("\n");
    expect(text).toMatch(/permanent/i);
  });
});

describe("draft happy path", () => {
  it("create_draft returns a sendable summary including the webLink", async () => {
    const created = {
      id: "draft-1",
      subject: "hello",
      to: ["me@example.com"],
      cc: [],
      bcc: [],
      body: "world",
      bodyContentType: "text" as const,
      webLink: "https://outlook.live.com/mail/0/AAA",
    };
    const { client } = await buildHarness({
      readOnly: false,
      enableOrganize: false,
      drafts: { async create() { return created as any; } },
    });
    const result = await client.callTool({
      name: "create_draft",
      arguments: { to: ["me@example.com"], subject: "hello", body: "world" },
    });
    const text = (result.content as any[]).map((c) => c.text).join("\n");
    expect(text).toContain("Draft saved");
    expect(text).toContain(created.webLink);
  });
});
