import { randomUUID } from "node:crypto";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export interface DeviceCodePrompt {
  url: string;
  code: string;
  message: string;
  expiresInSec: number;
}

export interface Prompter {
  /**
   * True if this prompter can hold the user-facing UI open while MSAL polls
   * to completion (e.g. an interactive terminal, an in-chat elicitation).
   * False if the prompter has no way to await the user (e.g. an MCP client
   * that doesn't advertise the elicitation capability) — in that case
   * TokenProvider throws AuthRequiredError immediately and lets MSAL keep
   * polling in the background.
   */
  readonly supportsInteractiveAwait: boolean;

  promptDeviceCode(prompt: DeviceCodePrompt): Promise<void>;
}

export class StderrPrompter implements Prompter {
  readonly supportsInteractiveAwait = true;

  async promptDeviceCode(prompt: DeviceCodePrompt): Promise<void> {
    process.stderr.write(`\n[outlook-mcp] ${prompt.message}\n\n`);
  }
}

export class McpElicitationPrompter implements Prompter {
  readonly supportsInteractiveAwait = true;

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
  readonly supportsInteractiveAwait = false;

  async promptDeviceCode(_prompt: DeviceCodePrompt): Promise<void> {
    // Intentionally a no-op. When supportsInteractiveAwait is false, the
    // TokenProvider throws AuthRequiredError directly with the URL+code; we
    // never invoke this method. It exists only to satisfy the interface.
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
