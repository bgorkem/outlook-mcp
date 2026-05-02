import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z, type ZodTypeAny } from "zod";
import { appendAudit } from "../security/audit.js";
import { GraphError } from "../graph/client.js";

export type ToolHandler<S extends ZodTypeAny> = (
  args: z.infer<S>,
) => Promise<CallToolResult> | CallToolResult;

interface ToolConfig<S extends ZodTypeAny> {
  description: string;
  inputSchema: S;
  annotations?: {
    title?: string;
    readOnlyHint?: boolean;
    destructiveHint?: boolean;
    idempotentHint?: boolean;
    openWorldHint?: boolean;
  };
}

export function registerAuditedTool<S extends ZodTypeAny>(
  server: McpServer,
  name: string,
  config: ToolConfig<S>,
  handler: ToolHandler<S>,
): void {
  server.registerTool(name, config as any, async (args: any) => {
    const start = Date.now();
    try {
      const result = await handler(args);
      appendAudit({
        tool: name,
        args,
        result: result.isError ? "error" : "ok",
        durationMs: Date.now() - start,
      });
      return result;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      appendAudit({
        tool: name,
        args,
        result: "error",
        durationMs: Date.now() - start,
        error: message,
      });
      return errorResult(err);
    }
  });
}

export function jsonResult(data: unknown, text?: string): CallToolResult {
  // MCP requires structuredContent to be an object, not a primitive or array.
  const structured: Record<string, unknown> = Array.isArray(data)
    ? { items: data, count: data.length }
    : data && typeof data === "object"
      ? (data as Record<string, unknown>)
      : { value: data };
  return {
    content: [{ type: "text", text: text ?? JSON.stringify(data, null, 2) }],
    structuredContent: structured,
  };
}

export function textResult(text: string): CallToolResult {
  return { content: [{ type: "text", text }] };
}

export function errorResult(err: unknown): CallToolResult {
  let msg: string;
  if (err instanceof GraphError) {
    msg = `Microsoft Graph error (${err.status}${err.code ? ` ${err.code}` : ""}): ${err.message}`;
  } else if (err instanceof Error) {
    msg = err.message;
  } else {
    msg = String(err);
  }
  return { content: [{ type: "text", text: msg }], isError: true };
}
