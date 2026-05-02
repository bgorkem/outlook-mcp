import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createPublicClient, resolveScopes, type ToolMode } from "./auth/msal.js";
import { TokenProvider } from "./auth/token.js";
import { GraphClient } from "./graph/client.js";
import { MessagesApi } from "./graph/messages.js";
import { DraftsApi } from "./graph/drafts.js";
import { registerReadTools } from "./tools/read.js";
import { registerDraftTools } from "./tools/drafts.js";
import { registerOrganizeTools } from "./tools/organize.js";

export interface ServerOptions {
  readOnly: boolean;
  enableOrganize: boolean;
}

export interface BuiltServer {
  server: McpServer;
  mode: ToolMode;
}

export function buildServer(opts: ServerOptions): BuiltServer {
  const mode: ToolMode = opts.readOnly ? "read-only" : "read-write";
  const pca = createPublicClient();
  const tokens = new TokenProvider({ pca, scopes: resolveScopes(mode) });
  const graph = new GraphClient(tokens);
  const messages = new MessagesApi(graph);
  const drafts = new DraftsApi(graph);

  const server = new McpServer({
    name: "outlook-mcp",
    version: "0.1.0",
  });

  registerReadTools(server, messages);

  if (!opts.readOnly) {
    registerDraftTools(server, drafts);
    if (opts.enableOrganize) {
      registerOrganizeTools(server, messages);
    }
  }

  return { server, mode };
}
