import type { DeviceCodeRequest, PublicClientApplication } from "@azure/msal-node";
import { AuthError } from "./errors.js";
import { parseMsalMessage, StderrPrompter, type Prompter } from "./prompter.js";

export { AuthError, AuthRequiredError } from "./errors.js";
export type AcquisitionMethod = "silent" | "device-code";

export interface TokenProviderOptions {
  pca: PublicClientApplication;
  scopes: string[];
  prompter?: Prompter;
}

export class TokenProvider {
  private inflight: Promise<string> | null = null;
  private readonly prompter: Prompter;
  lastAcquisitionMethod: AcquisitionMethod | null = null;

  constructor(private readonly opts: TokenProviderOptions) {
    this.prompter = opts.prompter ?? new StderrPrompter();
  }

  async getAccessToken(): Promise<string> {
    if (this.inflight) return this.inflight;
    this.inflight = this.acquire().finally(() => {
      this.inflight = null;
    });
    return this.inflight;
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
        process.stderr.write(
          `[outlook-mcp] silent token acquisition failed (${(err as Error).message}); falling back to device code\n`,
        );
      }
    }

    let promptError: Error | null = null;
    const request: DeviceCodeRequest = {
      scopes,
      deviceCodeCallback: (response) => {
        const prompt = parseMsalMessage(response.message, response.expiresIn);
        // Fire the prompter; MSAL keeps polling regardless of when the prompt's
        // promise settles. If the prompter rejects (e.g. NotSupportedPrompter),
        // cancel the device-code wait so we surface the error promptly.
        Promise.resolve()
          .then(() => this.prompter.promptDeviceCode(prompt))
          .catch((err: Error) => {
            promptError = err;
            request.cancel = true;
          });
      },
    };

    const result = await pca.acquireTokenByDeviceCode(request);

    if (promptError) throw promptError;
    if (!result?.accessToken) {
      throw new AuthError("Device code flow returned no access token");
    }
    this.lastAcquisitionMethod = "device-code";
    return result.accessToken;
  }
}
