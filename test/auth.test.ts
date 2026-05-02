import { describe, it, expect, vi } from "vitest";
import {
  StderrPrompter,
  McpElicitationPrompter,
  NotSupportedPrompter,
  parseMsalMessage,
  type Prompter,
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

describe("Prompter contracts", () => {
  it("StderrPrompter declares supportsInteractiveAwait=true and writes to stderr", async () => {
    const prompter = new StderrPrompter();
    expect(prompter.supportsInteractiveAwait).toBe(true);
    const spy = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    await prompter.promptDeviceCode({
      url: "https://example/link",
      code: "ABC12345",
      message: "go to https://example/link and enter ABC12345",
      expiresInSec: 900,
    });
    expect(spy).toHaveBeenCalledOnce();
    expect(String(spy.mock.calls[0]?.[0])).toContain("ABC12345");
    spy.mockRestore();
  });

  it("McpElicitationPrompter declares supportsInteractiveAwait=true and elicits with mode=url", async () => {
    const elicitInput = vi.fn().mockResolvedValue({ action: "accept" });
    const fakeServer = { server: { elicitInput } } as any;
    const prompter = new McpElicitationPrompter(fakeServer);
    expect(prompter.supportsInteractiveAwait).toBe(true);
    await prompter.promptDeviceCode({
      url: "https://www.microsoft.com/link",
      code: "AB1C2D3E",
      message: "ignored",
      expiresInSec: 900,
    });
    const arg = elicitInput.mock.calls[0]?.[0] as any;
    expect(arg.mode).toBe("url");
    expect(arg.url).toBe("https://www.microsoft.com/link");
    expect(arg.message).toContain("AB1C2D3E");
    expect(arg.elicitationId).toMatch(/^[0-9a-f-]{36}$/);
  });

  it("McpElicitationPrompter swallows elicitation errors so polling can continue", async () => {
    const elicitInput = vi.fn().mockRejectedValue(new Error("client decline"));
    const prompter = new McpElicitationPrompter({ server: { elicitInput } } as any);
    const spy = vi.spyOn(process.stderr, "write").mockImplementation(() => true);
    await expect(
      prompter.promptDeviceCode({ url: "u", code: "c", message: "m", expiresInSec: 0 }),
    ).resolves.toBeUndefined();
    spy.mockRestore();
  });

  it("NotSupportedPrompter declares supportsInteractiveAwait=false and is a no-op", async () => {
    const prompter = new NotSupportedPrompter();
    expect(prompter.supportsInteractiveAwait).toBe(false);
    await expect(
      prompter.promptDeviceCode({ url: "u", code: "c", message: "m", expiresInSec: 0 }),
    ).resolves.toBeUndefined();
  });
});

interface FakePcaOpts {
  cachedAccount?: boolean;
  silentToken?: string | null;
  deviceCodeMessage?: string;
  deviceCodeToken?: string | null;
  /** Hold acquireTokenByDeviceCode until this resolves, simulating polling. */
  pollingHoldTime?: number;
  pollingError?: Error;
}

function fakePca(opts: FakePcaOpts) {
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
      // MSAL fires the callback shortly after starting
      req.deviceCodeCallback({ message, expiresIn: 900 });
      // Then "polls" — caller may opt into a longer wait to simulate user inaction
      await new Promise((r) => setTimeout(r, opts.pollingHoldTime ?? 10));
      if (opts.pollingError) throw opts.pollingError;
      return opts.deviceCodeToken ? { accessToken: opts.deviceCodeToken } : null;
    },
  } as any;
}

const interactivePrompter = (spy: ReturnType<typeof vi.fn>): Prompter => ({
  supportsInteractiveAwait: true,
  promptDeviceCode: spy,
});

describe("TokenProvider — silent path", () => {
  it("returns a silently-acquired token without invoking the prompter", async () => {
    const promptSpy = vi.fn();
    const tokens = new TokenProvider({
      pca: fakePca({ cachedAccount: true, silentToken: "silent-token" }),
      scopes: ["Mail.Read"],
      prompter: interactivePrompter(promptSpy),
    });
    expect(await tokens.getAccessToken()).toBe("silent-token");
    expect(tokens.lastAcquisitionMethod).toBe("silent");
    expect(promptSpy).not.toHaveBeenCalled();
  });

  it("deduplicates concurrent silent acquisitions", async () => {
    let calls = 0;
    const pca = {
      getTokenCache: () => ({ getAllAccounts: async () => [{ homeAccountId: "x" }] }),
      acquireTokenSilent: async () => {
        calls++;
        return { accessToken: "silent-shared" };
      },
      acquireTokenByDeviceCode: async () => null,
    } as any;
    const tokens = new TokenProvider({
      pca,
      scopes: ["Mail.Read"],
      prompter: interactivePrompter(vi.fn()),
    });
    const [a, b, c] = await Promise.all([
      tokens.getAccessToken(),
      tokens.getAccessToken(),
      tokens.getAccessToken(),
    ]);
    expect([a, b, c]).toEqual(["silent-shared", "silent-shared", "silent-shared"]);
    expect(calls).toBe(1);
  });
});

describe("TokenProvider — interactive prompter path (StderrPrompter / elicitation)", () => {
  it("falls back to device code, surfaces the prompt, and awaits the token", async () => {
    const promptSpy = vi.fn().mockResolvedValue(undefined);
    const tokens = new TokenProvider({
      pca: fakePca({ deviceCodeToken: "dc-token" }),
      scopes: ["Mail.Read"],
      prompter: interactivePrompter(promptSpy),
    });
    expect(await tokens.getAccessToken()).toBe("dc-token");
    expect(tokens.lastAcquisitionMethod).toBe("device-code");
    expect(promptSpy).toHaveBeenCalledOnce();
    const arg = promptSpy.mock.calls[0]?.[0];
    expect(arg.code).toBe("AB1C2D3E");
    expect(arg.url).toBe("https://www.microsoft.com/link");
  });

  it("does not cancel polling if the prompter throws — MSAL still completes", async () => {
    // The interactive path is fire-and-forget on the prompter; an exception
    // from the prompter must not interrupt MSAL.
    const promptSpy = vi.fn().mockRejectedValue(new Error("rendering failed"));
    const tokens = new TokenProvider({
      pca: fakePca({ deviceCodeToken: "dc-token" }),
      scopes: ["Mail.Read"],
      prompter: interactivePrompter(promptSpy),
    });
    expect(await tokens.getAccessToken()).toBe("dc-token");
  });
});

describe("TokenProvider — non-interactive (NotSupported) path with background polling", () => {
  it("throws AuthRequiredError immediately with URL+code embedded so user can act from chat", async () => {
    const tokens = new TokenProvider({
      pca: fakePca({ pollingHoldTime: 1000, deviceCodeToken: "would-be-token" }),
      scopes: ["Mail.Read"],
      prompter: new NotSupportedPrompter(),
    });
    let caught: AuthRequiredError | null = null;
    try {
      await tokens.getAccessToken();
    } catch (e) {
      caught = e as AuthRequiredError;
    }
    expect(caught).toBeInstanceOf(AuthRequiredError);
    expect(caught!.message).toContain("https://www.microsoft.com/link");
    expect(caught!.message).toContain("AB1C2D3E");
    expect(caught!.message).toMatch(/sign[- ]in/i);
    expect(caught!.message).toMatch(/background/i);
  });

  it("a second call while polling is in progress reuses the SAME prompt (does not start a new code)", async () => {
    let callbackInvocations = 0;
    const pca = {
      getTokenCache: () => ({ getAllAccounts: async () => [] }),
      acquireTokenSilent: async () => null,
      acquireTokenByDeviceCode: async (req: any) => {
        callbackInvocations++;
        req.deviceCodeCallback({
          message:
            "To sign in, use https://www.microsoft.com/link and enter the code STABLE99 to authenticate.",
          expiresIn: 900,
        });
        await new Promise((r) => setTimeout(r, 200)); // simulate long polling
        return null;
      },
    } as any;

    const tokens = new TokenProvider({
      pca,
      scopes: ["Mail.Read"],
      prompter: new NotSupportedPrompter(),
    });

    // First call kicks off polling and throws
    await expect(tokens.getAccessToken()).rejects.toBeInstanceOf(AuthRequiredError);
    // Second call comes in BEFORE the simulated polling completes
    let second: AuthRequiredError | null = null;
    try {
      await tokens.getAccessToken();
    } catch (e) {
      second = e as AuthRequiredError;
    }
    expect(second).toBeInstanceOf(AuthRequiredError);
    expect(second!.message).toContain("STABLE99"); // same code, not a new one
    expect(callbackInvocations).toBe(1); // only one device-code request was issued
  });

  it("after background polling completes, the provider uses the silent path on next call", async () => {
    let backgroundResolve: ((token: string | null) => void) | null = null;
    let silentReady = false;
    const pca = {
      getTokenCache: () => ({
        getAllAccounts: async () => (silentReady ? [{ homeAccountId: "x" }] : []),
      }),
      acquireTokenSilent: async () =>
        silentReady ? { accessToken: "silent-after-bg" } : null,
      acquireTokenByDeviceCode: async (req: any) => {
        req.deviceCodeCallback({
          message:
            "To sign in, use https://www.microsoft.com/link and enter the code RESOLVE12 to authenticate.",
          expiresIn: 900,
        });
        return new Promise((resolve) => {
          backgroundResolve = (token) => {
            silentReady = true;
            resolve(token ? { accessToken: token } : null);
          };
        });
      },
    } as any;

    const tokens = new TokenProvider({
      pca,
      scopes: ["Mail.Read"],
      prompter: new NotSupportedPrompter(),
    });

    // First call: AuthRequiredError, background polling starts
    await expect(tokens.getAccessToken()).rejects.toBeInstanceOf(AuthRequiredError);

    // Simulate the user completing sign-in in their browser
    backgroundResolve!("bg-success-token");
    // Let the background promise's .then/.finally settle
    await new Promise((r) => setTimeout(r, 5));

    // Second call: silent path now hits the populated cache
    expect(await tokens.getAccessToken()).toBe("silent-after-bg");
  });

  it("when background polling fails (e.g., user never signed in), a retry starts a fresh device-code flow", async () => {
    let invocation = 0;
    const codes = ["FIRSTCODE", "SECONDCODE"];
    const pca = {
      getTokenCache: () => ({ getAllAccounts: async () => [] }),
      acquireTokenSilent: async () => null,
      acquireTokenByDeviceCode: async (req: any) => {
        const code = codes[invocation++] ?? "EXTRA";
        req.deviceCodeCallback({
          message: `Open https://www.microsoft.com/link and enter the code ${code} to authenticate.`,
          expiresIn: 900,
        });
        // First polling attempt fails; second succeeds
        if (invocation === 1) {
          await new Promise((r) => setTimeout(r, 5));
          throw new Error("expired_token");
        }
        return null;
      },
    } as any;

    const tokens = new TokenProvider({
      pca,
      scopes: ["Mail.Read"],
      prompter: new NotSupportedPrompter(),
    });

    let first: AuthRequiredError | null = null;
    try {
      await tokens.getAccessToken();
    } catch (e) {
      first = e as AuthRequiredError;
    }
    expect(first!.message).toContain("FIRSTCODE");

    // Wait for background polling to fail and clear in-flight state
    await new Promise((r) => setTimeout(r, 20));

    let second: AuthRequiredError | null = null;
    try {
      await tokens.getAccessToken();
    } catch (e) {
      second = e as AuthRequiredError;
    }
    expect(second!.message).toContain("SECONDCODE");
    expect(invocation).toBe(2); // a fresh flow was started after the first failed
  });
});
