// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import "isomorphic-fetch";
import { Client, type AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";

const authProvider: AuthenticationProvider = {
  getAccessToken: () => getAccessToken(),
};

export const graph = Client.initWithMiddleware({
  authProvider,
  defaultVersion: "v1.0",
});

export async function getMe(): Promise<{
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  id?: string;
}> {
  return await graph.api("/me").get();
}
