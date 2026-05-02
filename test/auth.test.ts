import { describe, it, expect, vi } from "vitest";
import {
  StderrPrompter,
  McpElicitationPrompter,
  NotSupportedPrompter,
  parseMsalMessage,
} from "../src/auth/prompter.js";
import { AuthRequiredError } from "../src/auth/errors.js";
import { TokenProvider } from "../src/auth/token.js";

describe("parseMsalMessage", () => {
  it("extracts URL and code from a real MSAL message", () => {
    const msg =
      "To sign in, use a web browser to open the page https://www.microsoft.com/link " +
      "and enter the code 2C7EYRH9 to authenticate.";
    const out = parseMsalMessage(msg, 900);
    expect(out.url).toBe("https://www.microsoft.com/link");
    expect(out.code).toBe("2C7EYRH9");
    expect(out.expiresInSec).toBe(900);
    expect(out.message).toBe(msg);
  });

  it("falls back to a sensible URL when parsing fails", () => {
    const out = parseMsalMessage("bizarre message no url", undefined);
    expect(out.url).toBe("https://www.microsoft.com/link");
    expect(out.expiresInSec).toBe(900);
  });
});

describe("StderrPrompter", () => {
  it("writes the MSAL message to stderr", async () => {
    const prompter = new StderrPrompter();
    const spy = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    await prompter.promptDeviceCode({
      url: "https://example/link",
      code: "ABC12345",
      message: "go to https://example/link and enter ABC12345",
      expiresInSec: 900,
    });
    expect(spy).toHaveBeenCalledOnce();
    const written = String(spy.mock.calls[0]?.[0]);
    expect(written).toContain("ABC12345");
    spy.mockRestore();
  });
});

describe("NotSupportedPrompter", () => {
  it("throws AuthRequiredError that embeds the actual URL and code so the user can act from chat", async () => {
    const prompter = new NotSupportedPrompter();
    await expect(
      prompter.promptDeviceCode({
        url: "https://www.microsoft.com/link",
        code: "AB1C2D3E",
        message: "msal raw msg",
        expiresInSec: 900,
      }),
    ).rejects.toBeInstanceOf(AuthRequiredError);

    try {
      await prompter.promptDeviceCode({
        url: "https://www.microsoft.com/link",
        code: "AB1C2D3E",
        message: "m",
        expiresInSec: 900,
      });
    } catch (e) {
      const err = e as AuthRequiredError;
      expect(err.message).toContain("https://www.microsoft.com/link");
      expect(err.message).toContain("AB1C2D3E");
      expect(err.message).toMatch(/sign[- ]in/i);
      // still mention the --login fallback for terminal users
      expect(err.message).toMatch(/--login/);
    }
  });
});

describe("McpElicitationPrompter", () => {
  it("calls server.server.elicitInput with mode=url and the device-code URL", async () => {
    const elicitInput = vi.fn().mockResolvedValue({ action: "accept" });
    const fakeServer = { server: { elicitInput } } as any;
    const prompter = new McpElicitationPrompter(fakeServer);
    await prompter.promptDeviceCode({
      url: "https://www.microsoft.com/link",
      code: "AB1C2D3E",
      message: "ignored",
      expiresInSec: 900,
    });
    expect(elicitInput).toHaveBeenCalledOnce();
    const arg = elicitInput.mock.calls[0]?.[0] as any;
    expect(arg.mode).toBe("url");
    expect(arg.url).toBe("https://www.microsoft.com/link");
    expect(arg.message).toContain("AB1C2D3E");
    expect(arg.elicitationId).toMatch(/^[0-9a-f-]{36}$/);
  });

  it("does not throw when elicitation fails — falls back to stderr log", async () => {
    const elicitInput = vi.fn().mockRejectedValue(new Error("client decline"));
    const fakeServer = { server: { elicitInput } } as any;
    const prompter = new McpElicitationPrompter(fakeServer);
    const spy = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    await expect(
      prompter.promptDeviceCode({ url: "u", code: "c", message: "m", expiresInSec: 0 }),
    ).resolves.toBeUndefined();
    spy.mockRestore();
  });
});

describe("TokenProvider", () => {
  function fakePca(opts: {
    cachedAccount?: boolean;
    silentToken?: string | null;
    deviceCodeMessage?: string;
    deviceCodeToken?: string | null;
    onCancel?: (req: any) => void;
  }) {
    return {
      getTokenCache: () => ({
        getAllAccounts: async () => (opts.cachedAccount ? [{ homeAccountId: "x" }] : []),
      }),
      acquireTokenSilent: async () =>
        opts.silentToken ? { accessToken: opts.silentToken } : null,
      acquireTokenByDeviceCode: async (req: any) => {
        const message =
          opts.deviceCodeMessage ??
          "To sign in, use https://www.microsoft.com/link and enter the code AB1C2D3E to authenticate.";
        // Simulate MSAL invoking the callback synchronously, then waiting one tick to "poll"
        req.deviceCodeCallback({ message, expiresIn: 900 });
        await new Promise((r) => setTimeout(r, 10));
        if (req.cancel) {
          opts.onCancel?.(req);
          return null;
        }
        return opts.deviceCodeToken ? { accessToken: opts.deviceCodeToken } : null;
      },
    } as any;
  }

  it("uses silent acquisition when a cached account exists", async () => {
    const promptSpy = vi.fn();
    const tokens = new TokenProvider({
      pca: fakePca({ cachedAccount: true, silentToken: "silent-token" }),
      scopes: ["Mail.Read"],
      prompter: { promptDeviceCode: promptSpy },
    });
    const token = await tokens.getAccessToken();
    expect(token).toBe("silent-token");
    expect(tokens.lastAcquisitionMethod).toBe("silent");
    expect(promptSpy).not.toHaveBeenCalled();
  });

  it("falls back to device code when no cached account; passes parsed prompt to prompter", async () => {
    const promptSpy = vi.fn().mockResolvedValue(undefined);
    const tokens = new TokenProvider({
      pca: fakePca({ deviceCodeToken: "dc-token" }),
      scopes: ["Mail.Read"],
      prompter: { promptDeviceCode: promptSpy },
    });
    const token = await tokens.getAccessToken();
    expect(token).toBe("dc-token");
    expect(tokens.lastAcquisitionMethod).toBe("device-code");
    expect(promptSpy).toHaveBeenCalledOnce();
    const arg = promptSpy.mock.calls[0]?.[0];
    expect(arg.code).toBe("AB1C2D3E");
    expect(arg.url).toBe("https://www.microsoft.com/link");
  });

  it("cancels device-code polling and rethrows when prompter rejects", async () => {
    const cancelSeen = vi.fn();
    const tokens = new TokenProvider({
      pca: fakePca({ deviceCodeToken: "should-not-be-returned", onCancel: cancelSeen }),
      scopes: ["Mail.Read"],
      prompter: new NotSupportedPrompter(),
    });
    await expect(tokens.getAccessToken()).rejects.toBeInstanceOf(AuthRequiredError);
    expect(cancelSeen).toHaveBeenCalled();
  });

  it("surfaces our AuthRequiredError even when MSAL throws device_code_polling_cancelled after cancel", async () => {
    // Real MSAL behaviour: when request.cancel = true is set during polling, MSAL
    // rejects acquireTokenByDeviceCode with a 'device_code_polling_cancelled' error
    // rather than resolving null. The TokenProvider must catch that and rethrow our
    // AuthRequiredError so the user sees the friendly message, not MSAL's raw error.
    const pca = {
      getTokenCache: () => ({ getAllAccounts: async () => [] }),
      acquireTokenSilent: async () => null,
      acquireTokenByDeviceCode: async (req: any) => {
        req.deviceCodeCallback({
          message: "go to https://www.microsoft.com/link and enter ZZZZ9999",
          expiresIn: 900,
        });
        await new Promise((r) => setTimeout(r, 5));
        // Mimic MSAL: throw, do not return null, on cancellation
        throw new Error(
          "device_code_polling_cancelled: Caller has cancelled token endpoint polling",
        );
      },
    } as any;

    const tokens = new TokenProvider({
      pca,
      scopes: ["Mail.Read"],
      prompter: new NotSupportedPrompter(),
    });
    await expect(tokens.getAccessToken()).rejects.toBeInstanceOf(AuthRequiredError);
  });

  it("deduplicates concurrent first-time acquisitions (one prompt for N callers)", async () => {
    const promptSpy = vi.fn().mockResolvedValue(undefined);
    const tokens = new TokenProvider({
      pca: fakePca({ deviceCodeToken: "shared-token" }),
      scopes: ["Mail.Read"],
      prompter: { promptDeviceCode: promptSpy },
    });
    const [a, b, c] = await Promise.all([
      tokens.getAccessToken(),
      tokens.getAccessToken(),
      tokens.getAccessToken(),
    ]);
    expect([a, b, c]).toEqual(["shared-token", "shared-token", "shared-token"]);
    expect(promptSpy).toHaveBeenCalledOnce();
  });
});
