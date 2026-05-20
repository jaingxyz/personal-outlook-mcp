import {
  PublicClientApplication,
  type Configuration,
  type TokenCacheContext,
  type ICachePlugin,
  LogLevel,
} from "@azure/msal-node";
import keytar from "keytar";
import { config } from "./config.js";

const keychainCachePlugin: ICachePlugin = {
  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    const data = await keytar.getPassword(
      config.keychainService,
      config.keychainAccount,
    );
    if (data) ctx.tokenCache.deserialize(data);
  },
  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (ctx.cacheHasChanged) {
      await keytar.setPassword(
        config.keychainService,
        config.keychainAccount,
        ctx.tokenCache.serialize(),
      );
    }
  },
};

const msalConfig: Configuration = {
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenant}`,
  },
  cache: {
    cachePlugin: keychainCachePlugin,
  },
  system: {
    loggerOptions: {
      // MSAL logs to stderr so we don't corrupt the MCP stdout stream.
      loggerCallback: (level, message) => {
        if (level <= LogLevel.Warning) console.error(`[msal] ${message}`);
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Warning,
    },
  },
};

const pca = new PublicClientApplication(msalConfig);

export class ReauthRequiredError extends Error {
  constructor(reason: string) {
    super(
      `Re-authentication required: ${reason}. Run \`npm run whoami\` from a terminal to refresh the token cache, then retry.`,
    );
    this.name = "ReauthRequiredError";
  }
}

export interface GetAccessTokenOptions {
  /**
   * If true, fall back to interactive device-code flow when silent refresh
   * fails. Use only from terminal-attached scripts (e.g. `npm run whoami`).
   * MCP tool calls run under Claude Desktop where stderr is invisible —
   * those must use the default (false) and surface ReauthRequiredError.
   */
  interactive?: boolean;
}

export async function getAccessToken(
  opts: GetAccessTokenOptions = {},
): Promise<string> {
  const tokenCache = pca.getTokenCache();
  const accounts = await tokenCache.getAllAccounts();

  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: config.scopes,
      });
      if (result?.accessToken) return result.accessToken;
    } catch (err) {
      if (!opts.interactive) {
        throw new ReauthRequiredError(silentFailureReason(err));
      }
      console.error(
        "[auth] silent token acquisition failed, falling back to device code:",
        err,
      );
    }
  } else if (!opts.interactive) {
    throw new ReauthRequiredError("no cached account");
  }

  const result = await pca.acquireTokenByDeviceCode({
    scopes: config.scopes,
    deviceCodeCallback: (resp) => {
      console.error("");
      console.error(resp.message);
      console.error("");
    },
  });

  if (!result?.accessToken) {
    throw new Error("Device code flow returned no access token");
  }
  return result.accessToken;
}

function silentFailureReason(err: unknown): string {
  if (err && typeof err === "object") {
    const e = err as { errorCode?: string; errorMessage?: string; message?: string };
    if (e.errorCode === "invalid_grant") {
      return "scopes changed or refresh token revoked";
    }
    if (e.errorCode) return e.errorCode;
    if (e.message) return e.message.slice(0, 200);
  }
  return "silent token acquisition failed";
}

export async function signOut(): Promise<void> {
  const tokenCache = pca.getTokenCache();
  const accounts = await tokenCache.getAllAccounts();
  for (const account of accounts) {
    await tokenCache.removeAccount(account);
  }
  await keytar.deletePassword(config.keychainService, config.keychainAccount);
}
