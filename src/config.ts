// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { readFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

function loadDotEnv(): void {
  const here = dirname(fileURLToPath(import.meta.url));
  const candidates = [
    resolve(here, "../../.env"),
    resolve(process.cwd(), ".env"),
  ];
  for (const path of candidates) {
    if (!existsSync(path)) continue;
    const text = readFileSync(path, "utf8");
    for (const rawLine of text.split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) continue;
      const eq = line.indexOf("=");
      if (eq === -1) continue;
      const key = line.slice(0, eq).trim();
      let val = line.slice(eq + 1).trim();
      if (
        (val.startsWith('"') && val.endsWith('"')) ||
        (val.startsWith("'") && val.endsWith("'"))
      ) {
        val = val.slice(1, -1);
      }
      if (process.env[key] === undefined) process.env[key] = val;
    }
    return;
  }
}

loadDotEnv();

function readVersion(): string {
  // package.json sits at the repo root, one level above dist/ (or src/ in
  // dev). Read it at runtime — a static `import` would pull a file outside
  // rootDir ("src") and break the tsc build / dist layout.
  const here = dirname(fileURLToPath(import.meta.url));
  const candidates = [
    resolve(here, "../package.json"),
    resolve(here, "../../package.json"),
  ];
  for (const path of candidates) {
    if (!existsSync(path)) continue;
    try {
      const pkg = JSON.parse(readFileSync(path, "utf8")) as {
        version?: string;
      };
      if (pkg.version) return pkg.version;
    } catch {
      // Malformed package.json — fall through to the next candidate.
    }
  }
  return "0.0.0";
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) {
    throw new Error(
      `Missing required env var ${name}. Copy .env.example to .env and fill it in.`,
    );
  }
  return v;
}

export const config = {
  version: readVersion(),
  clientId: required("AZURE_CLIENT_ID"),
  tenant: process.env.AZURE_TENANT || "consumers",
  scopes: [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "User.Read",
    "offline_access",
  ],
  defaultTimeZone: process.env.PERSONAL_OUTLOOK_TZ || "America/Los_Angeles",
  keychainService: "personal-outlook-mcp",
  keychainAccount: "msal-cache",
};
