import type { PublicClientApplication } from "@azure/msal-node";

export class AuthError extends Error {}

export interface TokenProviderOptions {
  pca: PublicClientApplication;
  scopes: string[];
}

export class TokenProvider {
  private inflight: Promise<string> | null = null;

  constructor(private readonly opts: TokenProviderOptions) {}

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
        if (silent?.accessToken) return silent.accessToken;
      } catch (err) {
        process.stderr.write(
          `[outlook-mcp] silent token acquisition failed (${(err as Error).message}); falling back to device code\n`,
        );
      }
    }

    const result = await pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        process.stderr.write(`\n[outlook-mcp] ${response.message}\n\n`);
      },
    });

    if (!result?.accessToken) {
      throw new AuthError("Device code flow returned no access token");
    }
    return result.accessToken;
  }
}
