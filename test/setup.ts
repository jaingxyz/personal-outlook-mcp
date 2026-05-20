// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz

// Tests import src/config indirectly through tool modules. config.ts requires
// AZURE_CLIENT_ID at module-load time. In CI there is no .env, so we have to
// set a placeholder before any source file is imported.
if (!process.env.AZURE_CLIENT_ID)
  process.env.AZURE_CLIENT_ID = "test-client-id";
if (!process.env.AZURE_TENANT) process.env.AZURE_TENANT = "consumers";
