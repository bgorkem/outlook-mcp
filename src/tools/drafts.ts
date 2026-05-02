import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { DraftsApi } from "../graph/drafts.js";
import { jsonResult, registerAuditedTool } from "./helpers.js";

const draftFields = {
  to: z.array(z.string().email()).optional(),
  cc: z.array(z.string().email()).optional(),
  bcc: z.array(z.string().email()).optional(),
  subject: z.string().max(998).optional(),
  body: z.string().optional(),
  bodyType: z.enum(["text", "html"]).optional(),
};

export function registerDraftTools(server: McpServer, drafts: DraftsApi): void {
  registerAuditedTool(
    server,
    "create_draft",
    {
      description:
        "Create a draft message in the user's Drafts folder. The draft is NEVER sent automatically — the user must click Send in Outlook. Optionally pass replyToMessageId (with replyAll) to create a reply draft seeded from an existing message.",
      inputSchema: z.object({
        ...draftFields,
        subject: z.string().max(998),
        body: z.string(),
        to: z.array(z.string().email()).min(1).optional(),
        replyToMessageId: z.string().optional(),
        replyAll: z.boolean().optional(),
      }),
      annotations: { destructiveHint: false, openWorldHint: true },
    },
    async (args) => {
      const draft = await drafts.create(args);
      return jsonResult(draft, summarize(draft));
    },
  );

  registerAuditedTool(
    server,
    "update_draft",
    {
      description:
        "Update an existing draft. Only provided fields are changed. Drafts are still never sent automatically.",
      inputSchema: z.object({
        id: z.string().min(1),
        ...draftFields,
      }),
      annotations: { destructiveHint: false, idempotentHint: true },
    },
    async ({ id, ...rest }) => {
      const draft = await drafts.update(id, rest);
      return jsonResult(draft, summarize(draft));
    },
  );

  registerAuditedTool(
    server,
    "list_drafts",
    {
      description: "List recent drafts in the Drafts folder.",
      inputSchema: z.object({ top: z.number().int().min(1).max(50).optional() }),
      annotations: { readOnlyHint: true, idempotentHint: true },
    },
    async ({ top }) => {
      const list = await drafts.list(top ?? 25);
      return jsonResult(list);
    },
  );

  registerAuditedTool(
    server,
    "delete_draft",
    {
      description:
        "Delete a draft by id. Refuses to operate on non-draft messages (use move_message to deleteditems for those).",
      inputSchema: z.object({ id: z.string().min(1) }),
      annotations: { destructiveHint: true, idempotentHint: false },
    },
    async ({ id }) => {
      await drafts.delete(id);
      return jsonResult({ id, deleted: true }, `Deleted draft ${id}`);
    },
  );
}

function summarize(d: { id: string; subject: string; to: string[]; webLink?: string }): string {
  return [
    `Draft saved (id: ${d.id}).`,
    `To: ${d.to.join(", ") || "(none)"}`,
    `Subject: ${d.subject}`,
    d.webLink ? `Open in Outlook to send: ${d.webLink}` : "",
  ]
    .filter(Boolean)
    .join("\n");
}
