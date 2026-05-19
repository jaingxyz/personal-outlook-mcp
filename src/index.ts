#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { buildServer } from "./server.js";

async function main(): Promise<void> {
  const server = buildServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // Server runs until the client closes stdin.
}

main().catch((err) => {
  console.error("[personal-outlook-mcp] fatal:", err);
  process.exit(1);
});
