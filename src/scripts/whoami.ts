// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { getAccessToken } from "../auth.js";
import { getMe } from "../graph.js";

async function main(): Promise<void> {
  // Force interactive flow if the keychain is missing or stale on the
  // scopes we need. Subsequent calls (incl. /me below) reuse the cached
  // token.
  await getAccessToken({ interactive: true });

  const me = await getMe();
  console.log(JSON.stringify(me, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
