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

export async function getAccessToken(): Promise<string> {
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
      console.error("[auth] silent token acquisition failed, falling back to device code:", err);
    }
  }

  const result = await pca.acquireTokenByDeviceCode({
    scopes: config.scopes,
    deviceCodeCallback: (resp) => {
      // Critical: write to stderr, never stdout.
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

export async function signOut(): Promise<void> {
  const tokenCache = pca.getTokenCache();
  const accounts = await tokenCache.getAllAccounts();
  for (const account of accounts) {
    await tokenCache.removeAccount(account);
  }
  await keytar.deletePassword(config.keychainService, config.keychainAccount);
}
