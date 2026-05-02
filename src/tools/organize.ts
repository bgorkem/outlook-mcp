import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { MessagesApi } from "../graph/messages.js";
import { jsonResult, registerAuditedTool } from "./helpers.js";

const FORBIDDEN_DESTINATIONS = new Set(["permanentdelete", "purge", "hardDelete"].map((s) => s.toLowerCase()));

const folderRef = z
  .string()
  .min(1)
  .refine((s) => !FORBIDDEN_DESTINATIONS.has(s.toLowerCase()), {
    message:
      "Permanent deletion is not supported by this server. Use 'deleteditems' to move the message to the Deleted Items folder instead.",
  });

export function registerOrganizeTools(server: McpServer, messages: MessagesApi): void {
  registerAuditedTool(
    server,
    "move_message",
    {
      description:
        "Move a message to another folder. destinationFolder accepts a well-known name (e.g., 'inbox', 'archive', 'deleteditems', 'junkemail') or a folder id. Permanent deletion is not supported.",
      inputSchema: z.object({
        id: z.string().min(1),
        destinationFolder: folderRef,
      }),
      annotations: { destructiveHint: true, idempotentHint: false },
    },
    async ({ id, destinationFolder }) => {
      const moved = await messages.moveMessage(id, destinationFolder);
      return jsonResult({ id: moved.id, destinationFolder }, `Moved message to ${destinationFolder} (new id: ${moved.id})`);
    },
  );

  registerAuditedTool(
    server,
    "mark_read",
    {
      description: "Mark a message as read.",
      inputSchema: z.object({ id: z.string().min(1) }),
      annotations: { destructiveHint: false, idempotentHint: true },
    },
    async ({ id }) => {
      await messages.setRead(id, true);
      return jsonResult({ id, isRead: true }, `Marked ${id} as read`);
    },
  );

  registerAuditedTool(
    server,
    "mark_unread",
    {
      description: "Mark a message as unread.",
      inputSchema: z.object({ id: z.string().min(1) }),
      annotations: { destructiveHint: false, idempotentHint: true },
    },
    async ({ id }) => {
      await messages.setRead(id, false);
      return jsonResult({ id, isRead: false }, `Marked ${id} as unread`);
    },
  );

  registerAuditedTool(
    server,
    "flag",
    {
      description: "Flag a message for follow-up.",
      inputSchema: z.object({ id: z.string().min(1) }),
      annotations: { destructiveHint: false, idempotentHint: true },
    },
    async ({ id }) => {
      await messages.setFlag(id, true);
      return jsonResult({ id, flagged: true }, `Flagged ${id}`);
    },
  );

  registerAuditedTool(
    server,
    "unflag",
    {
      description: "Remove the follow-up flag from a message.",
      inputSchema: z.object({ id: z.string().min(1) }),
      annotations: { destructiveHint: false, idempotentHint: true },
    },
    async ({ id }) => {
      await messages.setFlag(id, false);
      return jsonResult({ id, flagged: false }, `Unflagged ${id}`);
    },
  );

  registerAuditedTool(
    server,
    "create_folder",
    {
      description: "Create a new mail folder, optionally under a parent folder.",
      inputSchema: z.object({
        name: z.string().min(1).max(255),
        parentFolder: z.string().optional(),
      }),
      annotations: { destructiveHint: true, idempotentHint: false },
    },
    async ({ name, parentFolder }) => {
      const folder = await messages.createFolder(name, parentFolder);
      return jsonResult(folder, `Created folder "${folder.displayName}" (id: ${folder.id})`);
    },
  );

  registerAuditedTool(
    server,
    "rename_folder",
    {
      description: "Rename an existing mail folder.",
      inputSchema: z.object({
        id: z.string().min(1),
        newName: z.string().min(1).max(255),
      }),
      annotations: { destructiveHint: true, idempotentHint: true },
    },
    async ({ id, newName }) => {
      const folder = await messages.renameFolder(id, newName);
      return jsonResult(folder, `Renamed folder to "${folder.displayName}"`);
    },
  );
}
