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
      exclude: ["src/index.ts", "src/scripts/**"],
      // Thresholds set ~5% below current measurements so routine refactors
      // don't fail CI but a real regression in coverage will. Bump these as
      // coverage improves; never lower without a reason in the commit.
      thresholds: {
        lines: 65,
        statements: 65,
        functions: 45,
        branches: 60,
      },
    },
  },
});
