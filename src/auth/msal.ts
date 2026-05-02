import { PublicClientApplication, type Configuration } from "@azure/msal-node";
import { FileCachePlugin } from "./cache-plugin.js";
import { cacheFilePath } from "../util/paths.js";

export interface MsalEnv {
  clientId: string;
  authority: string;
}

export function readEnv(): MsalEnv {
  const clientId = process.env.OUTLOOK_MCP_CLIENT_ID?.trim();
  if (!clientId) {
    throw new Error(
      "OUTLOOK_MCP_CLIENT_ID env var is required. " +
        "Register an Azure AD app (multi-tenant + personal MS accounts, public client) " +
        "and pass its Application (client) ID. See README.md.",
    );
  }
  const authority =
    process.env.OUTLOOK_MCP_AUTHORITY?.trim() || "https://login.microsoftonline.com/consumers";
  return { clientId, authority };
}

export function createPublicClient(env: MsalEnv = readEnv()): PublicClientApplication {
  const config: Configuration = {
    auth: {
      clientId: env.clientId,
      authority: env.authority,
    },
    cache: {
      cachePlugin: new FileCachePlugin(cacheFilePath()),
    },
  };
  return new PublicClientApplication(config);
}

export type ToolMode = "read-only" | "read-write";

export function resolveScopes(mode: ToolMode): string[] {
  const base = mode === "read-only" ? ["Mail.Read"] : ["Mail.ReadWrite"];
  return [...base, "offline_access"];
}
