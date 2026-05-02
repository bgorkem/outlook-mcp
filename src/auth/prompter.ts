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
      const reason = err instanceof Error ? err.message : String(err);
      process.stderr.write(
        `[outlook-mcp] elicitation failed (${reason}); ` +
          `falling back to stderr prompt:\n${prompt.message}\n`,
      );
    }
  }
}

export class NotSupportedPrompter implements Prompter {
  async promptDeviceCode(prompt: DeviceCodePrompt): Promise<void> {
    const expiresMin = Math.max(1, Math.round(prompt.expiresInSec / 60));
    throw new AuthRequiredError(
      "**Outlook sign-in required.**\n\n" +
        `1. Open ${prompt.url} in your browser.\n` +
        `2. Enter this code: **${prompt.code}**\n` +
        `3. Sign in with your Outlook.com / Hotmail / Live account and approve the consent screen.\n` +
        "4. After the browser shows 'You can close this tab', ask me to retry.\n\n" +
        `(Code expires in ~${expiresMin} min. ` +
        "If your MCP client supports the `elicitation` capability, this prompt would appear in-chat instead — most clients don't yet.)\n\n" +
        "Alternative: run `npx -y @bgorkem/outlook-mcp --login` in a terminal once, then retry.",
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
