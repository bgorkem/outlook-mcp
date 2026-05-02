import { randomUUID } from "node:crypto";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { AuthRequiredError } from "./errors.js";

export interface DeviceCodePrompt {
  url: string;
  code: string;
  message: string;
  expiresInSec: number;
}

export interface Prompter {
  promptDeviceCode(prompt: DeviceCodePrompt): Promise<void>;
}

export class StderrPrompter implements Prompter {
  async promptDeviceCode(prompt: DeviceCodePrompt): Promise<void> {
    process.stderr.write(`\n[outlook-mcp] ${prompt.message}\n\n`);
  }
}

export class McpElicitationPrompter implements Prompter {
  constructor(private readonly server: McpServer) {}

  async promptDeviceCode(prompt: DeviceCodePrompt): Promise<void> {
    try {
      await this.server.server.elicitInput({
        mode: "url",
        url: prompt.url,
        message:
          `Sign in to your Microsoft account to grant outlook-mcp access. ` +
          `Open the link, then enter the code ${prompt.code} when prompted. ` +
          `(Code expires in ${Math.round(prompt.expiresInSec / 60)} min.)`,
        elicitationId: randomUUID(),
      });
    } catch (err) {
      process.stderr.write(
        `[outlook-mcp] elicitation failed (${(err as Error).message}); ` +
          `falling back to stderr prompt:\n${prompt.message}\n`,
      );
    }
  }
}

export class NotSupportedPrompter implements Prompter {
  async promptDeviceCode(_prompt: DeviceCodePrompt): Promise<void> {
    throw new AuthRequiredError(
      "Outlook authentication required.\n" +
        "This MCP client does not support in-session sign-in prompts " +
        "(it did not advertise the 'elicitation' capability).\n\n" +
        "Run this once in a terminal, then retry your request:\n" +
        "    npx -y @bgorkem/outlook-mcp --login\n\n" +
        "Make sure OUTLOOK_MCP_CLIENT_ID is set in the same shell.",
    );
  }
}

const DEVICE_CODE_RE = /enter the code\s+([A-Za-z0-9-]+)/i;
const URL_RE = /(https?:\/\/[^\s]+)/;

export function parseMsalMessage(message: string, expiresInSec?: number): DeviceCodePrompt {
  const codeMatch = message.match(DEVICE_CODE_RE);
  const urlMatch = message.match(URL_RE);
  return {
    url: urlMatch?.[1] ?? "https://www.microsoft.com/link",
    code: codeMatch?.[1] ?? "(see message)",
    message,
    expiresInSec: expiresInSec ?? 900,
  };
}
