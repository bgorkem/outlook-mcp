import type { AuthenticationResult, PublicClientApplication } from "@azure/msal-node";
import { AuthError, AuthRequiredError } from "./errors.js";
import {
  parseMsalMessage,
  StderrPrompter,
  type DeviceCodePrompt,
  type Prompter,
} from "./prompter.js";

export { AuthError, AuthRequiredError } from "./errors.js";
export type AcquisitionMethod = "silent" | "device-code";

export interface TokenProviderOptions {
  pca: PublicClientApplication;
  scopes: string[];
  prompter?: Prompter;
}

export class TokenProvider {
  private acquireInflight: Promise<string> | null = null;
  private backgroundFlow: Promise<AuthenticationResult | null> | null = null;
  private lastPrompt: DeviceCodePrompt | null = null;
  private readonly prompter: Prompter;
  lastAcquisitionMethod: AcquisitionMethod | null = null;

  constructor(private readonly opts: TokenProviderOptions) {
    this.prompter = opts.prompter ?? new StderrPrompter();
  }

  async getAccessToken(): Promise<string> {
    if (this.acquireInflight) return this.acquireInflight;
    this.acquireInflight = this.acquire().finally(() => {
      this.acquireInflight = null;
    });
    return this.acquireInflight;
  }

  private async acquire(): Promise<string> {
    const { pca, scopes } = this.opts;

    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      try {
        const account = accounts[0]!;
        const silent = await pca.acquireTokenSilent({ account, scopes });
        if (silent?.accessToken) {
          this.lastAcquisitionMethod = "silent";
          return silent.accessToken;
        }
      } catch (err) {
        const reason = err instanceof Error ? err.message : String(err);
        process.stderr.write(
          `[outlook-mcp] silent token acquisition failed (${reason}); falling back to device code\n`,
        );
      }
    }

    // No silent token. Two strategies depending on whether the prompter can
    // hold a UI open while we await MSAL polling:
    //   - supportsInteractiveAwait: classic flow — prompter surfaces the
    //     URL+code (e.g. via stderr for --login, or in-chat elicitation for
    //     capable MCP clients), and we await acquireTokenByDeviceCode.
    //   - !supportsInteractiveAwait: we cannot hold the tool call open for
    //     ~15 minutes waiting for the user to act, so we throw an
    //     AuthRequiredError immediately with the URL+code embedded, and let
    //     MSAL keep polling in the background. The cache plugin writes
    //     ~/.outlook-mcp/cache.json once the user signs in; subsequent tool
    //     calls go through the silent path above and succeed.
    return this.prompter.supportsInteractiveAwait
      ? this.acquireAwaitingPrompter()
      : this.acquireWithBackgroundPolling();
  }

  private async acquireAwaitingPrompter(): Promise<string> {
    const { pca, scopes } = this.opts;
    const result = await pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        const prompt = parseMsalMessage(response.message, response.expiresIn);
        // Fire-and-forget — MSAL polls regardless of when the prompter resolves.
        Promise.resolve()
          .then(() => this.prompter.promptDeviceCode(prompt))
          .catch((err) => {
            const reason = err instanceof Error ? err.message : String(err);
            process.stderr.write(`[outlook-mcp] prompter error: ${reason}\n`);
          });
      },
    });
    if (!result?.accessToken) {
      throw new AuthError("Device code flow returned no access token");
    }
    this.lastAcquisitionMethod = "device-code";
    return result.accessToken;
  }

  private async acquireWithBackgroundPolling(): Promise<string> {
    const { pca, scopes } = this.opts;

    // If a flow is already in progress, surface the SAME prompt rather than
    // starting a second concurrent flow with a different code.
    if (this.backgroundFlow) {
      if (this.lastPrompt) throw this.makeAuthRequiredError(this.lastPrompt);
      throw new AuthRequiredError(
        "Outlook sign-in is in progress. Please wait a moment and try again.",
      );
    }

    let promptResolve!: (p: DeviceCodePrompt) => void;
    let promptReject!: (err: unknown) => void;
    const promptReady = new Promise<DeviceCodePrompt>((res, rej) => {
      promptResolve = res;
      promptReject = rej;
    });

    const flow = pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        const prompt = parseMsalMessage(response.message, response.expiresIn);
        this.lastPrompt = prompt;
        promptResolve(prompt);
      },
    });

    this.backgroundFlow = flow;

    flow
      .then((result) => {
        if (result?.accessToken) {
          this.lastAcquisitionMethod = "device-code";
          process.stderr.write(
            "[outlook-mcp] background sign-in completed; token cache updated\n",
          );
        } else {
          process.stderr.write(
            "[outlook-mcp] background sign-in returned no token (likely expired or declined)\n",
          );
        }
      })
      .catch((err) => {
        // If MSAL rejected before deviceCodeCallback fired (early misconfiguration,
        // network failure, etc.), unblock the awaiter so the tool call returns the
        // real MSAL error instead of hanging on promptReady forever.
        // Once promptReady is already settled, this reject is a no-op.
        promptReject(err);
        const reason = err instanceof Error ? err.message : String(err);
        process.stderr.write(`[outlook-mcp] background sign-in failed: ${reason}\n`);
      })
      .finally(() => {
        this.backgroundFlow = null;
        this.lastPrompt = null;
      });

    const prompt = await promptReady;
    throw this.makeAuthRequiredError(prompt);
  }

  private makeAuthRequiredError(prompt: DeviceCodePrompt): AuthRequiredError {
    const expiresMin = Math.max(1, Math.round(prompt.expiresInSec / 60));
    return new AuthRequiredError(
      "**Outlook sign-in required.**\n\n" +
        `1. Open ${prompt.url} in your browser.\n` +
        `2. Enter this code: **${prompt.code}**\n` +
        "3. Sign in with your Outlook.com / Hotmail / Live account and approve the consent screen.\n" +
        "4. Once the browser shows 'You can close this tab', ask me to retry.\n\n" +
        `(Code expires in ~${expiresMin} min. ` +
        "I'm polling Microsoft in the background — once you sign in, I'll cache the token automatically. " +
        "If your MCP client supports the `elicitation` capability, this prompt would appear inline; most don't yet.)",
    );
  }
}
