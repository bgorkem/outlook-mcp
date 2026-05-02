import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createPublicClient, resolveScopes, type ToolMode } from "./auth/msal.js";
import { TokenProvider } from "./auth/token.js";
import {
  McpElicitationPrompter,
  NotSupportedPrompter,
  type Prompter,
} from "./auth/prompter.js";
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
  const scopes = resolveScopes(mode);

  const server = new McpServer({
    name: "outlook-mcp",
    version: "0.1.2",
  });

  // Auth is wired lazily because client capabilities (specifically `elicitation`)
  // aren't known until the client has completed `initialize`. The first tool call
  // arrives after that, so resolving the prompter inside getTokens() is safe.
  let tokens: TokenProvider | null = null;
  const getTokens = (): TokenProvider => {
    if (tokens) return tokens;
    const prompter = pickPrompter(server);
    tokens = new TokenProvider({ pca, scopes, prompter });
    return tokens;
  };

  const graph = new GraphClient({ getAccessToken: () => getTokens().getAccessToken() });
  const messages = new MessagesApi(graph);
  const drafts = new DraftsApi(graph);

  registerReadTools(server, messages);

  if (!opts.readOnly) {
    registerDraftTools(server, drafts);
    if (opts.enableOrganize) {
      registerOrganizeTools(server, messages);
    }
  }

  return { server, mode };
}

function pickPrompter(server: McpServer): Prompter {
  const caps = server.server.getClientCapabilities();
  if (caps?.elicitation) {
    return new McpElicitationPrompter(server);
  }
  return new NotSupportedPrompter();
}
