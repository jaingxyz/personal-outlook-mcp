// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { describe, it, expect, vi, beforeEach } from "vitest";

// auth.ts imports config which validates AZURE_CLIENT_ID; provide one before importing.
process.env.AZURE_CLIENT_ID = "test-client-id";
process.env.AZURE_TENANT = "consumers";

const tokenCacheState = {
  accounts: [] as Array<{ homeAccountId: string }>,
  acquireTokenSilentImpl: vi.fn(),
  acquireTokenByDeviceCodeImpl: vi.fn(),
};

vi.mock("@azure/msal-node", () => {
  return {
    LogLevel: { Warning: 2, Error: 0, Info: 3, Verbose: 4 },
    PublicClientApplication: class {
      getTokenCache() {
        return {
          getAllAccounts: async () => tokenCacheState.accounts,
          removeAccount: async () => undefined,
          serialize: () => "",
          deserialize: () => undefined,
        };
      }
      acquireTokenSilent(opts: unknown) {
        return tokenCacheState.acquireTokenSilentImpl(opts);
      }
      acquireTokenByDeviceCode(opts: unknown) {
        return tokenCacheState.acquireTokenByDeviceCodeImpl(opts);
      }
    },
  };
});

vi.mock("keytar", () => ({
  default: {
    getPassword: async () => null,
    setPassword: async () => undefined,
    deletePassword: async () => undefined,
  },
}));

beforeEach(() => {
  tokenCacheState.accounts = [];
  tokenCacheState.acquireTokenSilentImpl.mockReset();
  tokenCacheState.acquireTokenByDeviceCodeImpl.mockReset();
});

describe("getAccessToken", () => {
  it("throws ReauthRequiredError when no account is cached and interactive=false", async () => {
    const { getAccessToken, ReauthRequiredError } =
      await import("../src/auth.js");
    tokenCacheState.accounts = [];
    await expect(getAccessToken({ interactive: false })).rejects.toBeInstanceOf(
      ReauthRequiredError,
    );
  });

  it("throws ReauthRequiredError when silent acquisition fails and interactive=false", async () => {
    const { getAccessToken, ReauthRequiredError } =
      await import("../src/auth.js");
    tokenCacheState.accounts = [{ homeAccountId: "x" }];
    tokenCacheState.acquireTokenSilentImpl.mockRejectedValue(
      Object.assign(new Error("invalid_grant"), { errorCode: "invalid_grant" }),
    );
    await expect(getAccessToken({ interactive: false })).rejects.toBeInstanceOf(
      ReauthRequiredError,
    );
    // Crucially: do not fall back to device code from non-interactive path.
    expect(tokenCacheState.acquireTokenByDeviceCodeImpl).not.toHaveBeenCalled();
  });

  it("falls back to device code when silent fails AND interactive=true", async () => {
    const { getAccessToken } = await import("../src/auth.js");
    tokenCacheState.accounts = [{ homeAccountId: "x" }];
    tokenCacheState.acquireTokenSilentImpl.mockRejectedValue(
      new Error("expired"),
    );
    tokenCacheState.acquireTokenByDeviceCodeImpl.mockResolvedValue({
      accessToken: "device-code-token",
    });
    const tok = await getAccessToken({ interactive: true });
    expect(tok).toBe("device-code-token");
  });

  it("returns silent token when account is cached and silent succeeds", async () => {
    const { getAccessToken } = await import("../src/auth.js");
    tokenCacheState.accounts = [{ homeAccountId: "x" }];
    tokenCacheState.acquireTokenSilentImpl.mockResolvedValue({
      accessToken: "silent-token",
    });
    expect(await getAccessToken({ interactive: false })).toBe("silent-token");
    expect(tokenCacheState.acquireTokenByDeviceCodeImpl).not.toHaveBeenCalled();
  });

  it("ReauthRequiredError message points the user at npm run whoami", async () => {
    const { getAccessToken, ReauthRequiredError } =
      await import("../src/auth.js");
    tokenCacheState.accounts = [];
    try {
      await getAccessToken({ interactive: false });
    } catch (e) {
      expect(e).toBeInstanceOf(ReauthRequiredError);
      expect((e as Error).message).toMatch(/npm run whoami/);
    }
  });
});
