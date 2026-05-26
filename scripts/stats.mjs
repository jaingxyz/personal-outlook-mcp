#!/usr/bin/env node
// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz

// Snapshot usage stats from npm, Smithery, and GitHub into a local JSONL.
// Run on demand or on a cron. Output is gitignored — strictly local.
//
// Usage:
//   node scripts/stats.mjs                  # appends one line to stats.jsonl
//   node scripts/stats.mjs --print          # also pretty-prints to stdout
//   node scripts/stats.mjs --history        # show all past snapshots, no fetch
//
// Why JSONL: each run is one line, append-only, easy to grep/jq/diff. No
// schema migrations needed when fields change.

import { appendFileSync, readFileSync, existsSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const outFile = resolve(repoRoot, "stats.jsonl");

const args = new Set(process.argv.slice(2));

if (args.has("--history")) {
  if (!existsSync(outFile)) {
    console.log(
      "(no history yet — run without --history to capture first snapshot)",
    );
    process.exit(0);
  }
  const lines = readFileSync(outFile, "utf8")
    .trim()
    .split("\n")
    .filter(Boolean);
  console.log(`# ${lines.length} snapshot(s)\n`);
  for (const line of lines) {
    const s = JSON.parse(line);
    console.log(
      `${s.timestamp.slice(0, 10)}  npm/7d=${s.npm.downloads7d}  views/14d=${s.github.views14d}  clones/14d=${s.github.clones14d}  stars=${s.github.stars}  smithery.useCount=${s.smithery.useCount}`,
    );
  }
  process.exit(0);
}

const PKG = "@jaingxyz/personal-outlook-mcp";
const REPO = "jaingxyz/personal-outlook-mcp";

const today = new Date();
const since = new Date(today.getTime() - 7 * 86400000);
const npmRange = `${since.toISOString().slice(0, 10)}:${today.toISOString().slice(0, 10)}`;

const snap = {
  timestamp: today.toISOString(),
  package: PKG,
  npm: { downloads7d: null, latestVersion: null, error: null },
  smithery: { useCount: null, score: null, verified: null, error: null },
  github: {
    views14d: null,
    viewsUnique14d: null,
    clones14d: null,
    clonesUnique14d: null,
    stars: null,
    forks: null,
    error: null,
  },
};

async function fetchJson(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`HTTP ${r.status} on ${url}`);
  return r.json();
}

// --- npm downloads
try {
  const d = await fetchJson(
    `https://api.npmjs.org/downloads/range/${npmRange}/${encodeURIComponent(PKG)}`,
  );
  snap.npm.downloads7d = (d.downloads || []).reduce(
    (a, x) => a + (x.downloads || 0),
    0,
  );
} catch (e) {
  snap.npm.error = String(e.message ?? e).slice(0, 200);
}

try {
  const r = await fetchJson(
    `https://registry.npmjs.org/${encodeURIComponent(PKG)}`,
  );
  snap.npm.latestVersion = r["dist-tags"]?.latest ?? null;
} catch (e) {
  snap.npm.error =
    (snap.npm.error ?? "") + " | " + String(e.message ?? e).slice(0, 100);
}

// --- Smithery
try {
  const r = await fetchJson(`https://api.smithery.ai/servers/${REPO}`);
  snap.smithery.useCount = r.useCount ?? 0;
  snap.smithery.verified = r.verified ?? false;
} catch (e) {
  snap.smithery.error = String(e.message ?? e).slice(0, 200);
}

try {
  const search = await fetchJson(
    `https://api.smithery.ai/servers?q=${encodeURIComponent(REPO)}&pageSize=5`,
  );
  for (const s of search.servers ?? []) {
    if (s.qualifiedName === REPO) {
      if (typeof s.score === "number") snap.smithery.score = s.score;
      break;
    }
  }
} catch {
  // search index miss is fine — score may not be exposed there
}

// --- GitHub (admin-only endpoints — needs `gh` CLI authenticated as repo admin)
function ghJson(path) {
  return JSON.parse(execFileSync("gh", ["api", path], { encoding: "utf8" }));
}

try {
  const v = ghJson(`repos/${REPO}/traffic/views`);
  snap.github.views14d = v.count ?? 0;
  snap.github.viewsUnique14d = v.uniques ?? 0;
} catch (e) {
  snap.github.error = String(e.message ?? e).slice(0, 200);
}

try {
  const c = ghJson(`repos/${REPO}/traffic/clones`);
  snap.github.clones14d = c.count ?? 0;
  snap.github.clonesUnique14d = c.uniques ?? 0;
} catch (e) {
  snap.github.error =
    (snap.github.error ?? "") + " | " + String(e.message ?? e).slice(0, 100);
}

try {
  const r = ghJson(`repos/${REPO}`);
  snap.github.stars = r.stargazers_count ?? 0;
  snap.github.forks = r.forks_count ?? 0;
} catch (e) {
  snap.github.error =
    (snap.github.error ?? "") + " | " + String(e.message ?? e).slice(0, 100);
}

appendFileSync(outFile, JSON.stringify(snap) + "\n");

if (args.has("--print")) {
  console.log(JSON.stringify(snap, null, 2));
} else {
  console.log(
    `[stats] ${snap.timestamp.slice(0, 10)}  npm/7d=${snap.npm.downloads7d}  views/14d=${snap.github.views14d}  clones/14d=${snap.github.clones14d}  stars=${snap.github.stars}  smithery.useCount=${snap.smithery.useCount}`,
  );
  console.log(`[stats] appended to ${outFile}`);
}
