#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { buildServer } from "./server.js";
import { createPublicClient, resolveScopes } from "./auth/msal.js";
import { TokenProvider } from "./auth/token.js";
import { cacheFilePath } from "./util/paths.js";

interface CliFlags {
  readOnly: boolean;
  enableOrganize: boolean;
  showHelp: boolean;
  showVersion: boolean;
  login: boolean;
}

function parseArgs(argv: string[]): CliFlags {
  const flags: CliFlags = {
    readOnly: false,
    enableOrganize: false,
    showHelp: false,
    showVersion: false,
    login: false,
  };
  for (const arg of argv) {
    switch (arg) {
      case "--read-only":
        flags.readOnly = true;
        break;
      case "--enable-organize":
        flags.enableOrganize = true;
        break;
      case "--login":
        flags.login = true;
        break;
      case "-h":
      case "--help":
        flags.showHelp = true;
        break;
      case "-v":
      case "--version":
        flags.showVersion = true;
        break;
      default:
        if (arg.startsWith("-")) {
          process.stderr.write(`[outlook-mcp] unknown flag: ${arg}\n`);
        }
    }
  }
  return flags;
}

function helpText(): string {
  return `outlook-mcp — MCP server for personal Microsoft mailboxes

USAGE
  outlook-mcp [--read-only] [--enable-organize]
  outlook-mcp --login                  # one-shot device-code sign-in, then exit

FLAGS
  --login             Run the device-code sign-in flow once and exit.
                      Use this for first-time auth or to refresh consent.
                      No MCP transport is started.
  --read-only         Register only read tools. Requests Mail.Read scope only.
  --enable-organize   Additionally register move/mark/flag/folder tools.
                      Without this flag, only read + draft tools are exposed.

ENV
  OUTLOOK_MCP_CLIENT_ID    (required)  Azure AD app (client) id
  OUTLOOK_MCP_AUTHORITY    (optional)  default: https://login.microsoftonline.com/consumers
  OUTLOOK_MCP_CACHE_DIR    (optional)  default: ~/.outlook-mcp

NOTES
  This server exposes NO send_mail tool by design. Drafts are created in
  your mailbox; you click Send in Outlook. The Mail.Send scope is never
  requested. Permanent deletion is also not supported.
`;
}

async function runLogin(readOnly: boolean): Promise<void> {
  const pca = createPublicClient();
  const scopes = resolveScopes(readOnly ? "read-only" : "read-write");
  const tokens = new TokenProvider({ pca, scopes });
  process.stderr.write(
    `[outlook-mcp] acquiring token (silent first; device code if needed). Scopes: ${scopes.join(", ")}\n`,
  );
  await tokens.getAccessToken();
  if (tokens.lastAcquisitionMethod === "silent") {
    process.stderr.write(
      `[outlook-mcp] silent acquisition succeeded — existing cache at ${cacheFilePath()} is valid.\n`,
    );
  } else {
    process.stderr.write(
      `[outlook-mcp] sign-in complete. Token cache location: ${cacheFilePath()}\n`,
    );
  }
  process.stderr.write(`[outlook-mcp] you can now wire this server to your MCP client.\n`);
}

async function main(): Promise<void> {
  const flags = parseArgs(process.argv.slice(2));

  if (flags.showHelp) {
    process.stdout.write(helpText());
    return;
  }
  if (flags.showVersion) {
    process.stdout.write("outlook-mcp 0.1.2\n");
    return;
  }
  if (flags.login) {
    await runLogin(flags.readOnly);
    return;
  }

  const { server, mode } = buildServer({
    readOnly: flags.readOnly,
    enableOrganize: flags.enableOrganize,
  });

  process.stderr.write(
    `[outlook-mcp] starting in ${mode} mode (organize tools: ${flags.enableOrganize ? "on" : "off"}). ` +
      `Send tools are NEVER registered.\n`,
  );

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  process.stderr.write(`[outlook-mcp] fatal: ${err instanceof Error ? err.message : String(err)}\n`);
  process.exit(1);
});
