// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz

// Integration test: spawns the compiled MCP server as a child process and
// exercises it over stdio with real JSON-RPC messages. No Graph mocking,
// no network calls — `tools/list` is auth-free and self-contained.
//
// Requires `npm run build` to have produced dist/. The test skips itself
// if dist/index.js is missing so a fresh checkout doesn't fail on a missing
// build artifact.

import { describe, it, expect } from "vitest";
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { resolve } from "node:path";

const DIST_PATH = resolve(__dirname, "..", "dist", "index.js");
const distExists = existsSync(DIST_PATH);

const describeIfBuilt = distExists ? describe : describe.skip;

describeIfBuilt("MCP server stdio protocol", () => {
  it("responds to initialize and tools/list with the expected tool surface", async () => {
    const responses = await runServer([
      {
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2024-11-05",
          capabilities: {},
          clientInfo: { name: "integration-test", version: "0" },
        },
      },
      { jsonrpc: "2.0", method: "notifications/initialized" },
      { jsonrpc: "2.0", id: 2, method: "tools/list" },
    ]);

    const init = responses.find((r) => r.id === 1);
    expect(init?.result?.serverInfo?.name).toBe("personal-outlook-mcp");
    expect(init?.result?.protocolVersion).toBe("2024-11-05");

    const list = responses.find((r) => r.id === 2);
    const tools = list?.result?.tools as Array<{
      name: string;
      inputSchema: unknown;
    }>;
    expect(Array.isArray(tools)).toBe(true);

    const names = tools.map((t) => t.name).sort();

    // Snapshot: every tool we expose, lexicographically sorted.
    expect(names).toEqual(
      [
        "personal_calendar_cancel_event",
        "personal_calendar_create_event",
        "personal_calendar_list_calendars",
        "personal_calendar_list_events",
        "personal_calendar_read_event",
        "personal_calendar_respond_to_invite",
        "personal_calendar_update_event",
        "personal_email_create_draft",
        "personal_email_delete",
        "personal_email_download_attachment",
        "personal_email_list_attachments",
        "personal_email_list_folders",
        "personal_email_list_recent",
        "personal_email_mark_read",
        "personal_email_move",
        "personal_email_read",
        "personal_email_reply",
        "personal_email_search",
        "personal_email_send",
        "personal_email_send_draft",
      ].sort(),
    );

    for (const t of tools) {
      expect(t.inputSchema).toBeTypeOf("object");
    }
  }, 15000);
});

interface RpcMessage {
  jsonrpc: string;
  id?: number;
  method?: string;
  params?: unknown;
  result?: any;
  error?: unknown;
}

async function runServer(messages: object[]): Promise<RpcMessage[]> {
  const child = spawn(process.execPath, [DIST_PATH], {
    env: {
      ...process.env,
      AZURE_CLIENT_ID: "integration-test-client-id",
      AZURE_TENANT: "consumers",
    },
    stdio: ["pipe", "pipe", "pipe"],
  });

  for (const m of messages) {
    child.stdin.write(JSON.stringify(m) + "\n");
  }

  let stdout = "";
  child.stdout.on("data", (chunk: Buffer) => {
    stdout += chunk.toString("utf8");
  });

  // Give the server time to respond, then close stdin to make it exit.
  await new Promise((r) => setTimeout(r, 1500));
  child.stdin.end();

  await new Promise<void>((resolve) => {
    child.on("exit", () => resolve());
    setTimeout(() => {
      child.kill();
      resolve();
    }, 5000);
  });

  return stdout
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line) as RpcMessage;
      } catch {
        return null;
      }
    })
    .filter((m): m is RpcMessage => m !== null);
}
