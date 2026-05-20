// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    setupFiles: ["./test/setup.ts"],
    coverage: {
      provider: "v8",
      reporter: ["text", "html"],
      include: ["src/**/*.ts"],
      // Exclude:
      //  - index.ts: trivial bootstrap, just wires server -> stdio.
      //  - scripts/: CLI tooling, not part of the MCP runtime.
      //  - config.ts: env loading + .env parsing. Coverage depends on
      //    whether .env exists at test time, which differs between local
      //    and CI; the logic is straightforward enough that the integration
      //    test (which loads config end-to-end) catches breakage.
      //  - graph.ts: a singleton wiring auth provider to the Graph SDK.
      //    Exercised by every tool test transitively but v8 reports it as
      //    50% because of init code that runs at import time only.
      exclude: [
        "src/index.ts",
        "src/scripts/**",
        "src/config.ts",
        "src/graph.ts",
        // server.ts is wiring (Zod schemas -> McpServer.registerTool). The
        // integration test spawns the compiled binary and asserts every
        // tool registers correctly, which is more meaningful than per-line
        // coverage of the registration calls themselves.
        "src/server.ts",
      ],
      // Thresholds set ~5% below current measurements so routine refactors
      // don't fail CI but a real regression in coverage will. Bump these as
      // coverage improves; never lower without a reason in the commit.
      thresholds: {
        lines: 75,
        statements: 75,
        functions: 70,
        branches: 55,
      },
    },
  },
});
