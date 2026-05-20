// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { vi } from "vitest";

export interface RecordedCall {
  path: string;
  method: "get" | "post" | "patch" | "delete";
  body?: unknown;
  headers: Record<string, string>;
  query?: Record<string, unknown>;
  select?: string;
  top?: number;
  filter?: string;
  orderby?: string;
  search?: string;
}

export interface MockGraph {
  api: ReturnType<typeof vi.fn>;
  calls: RecordedCall[];
  /** Queue of responses keyed by HTTP method; pop in order. */
  responses: { method: string; value: unknown }[];
}

export function makeMockGraph(): MockGraph {
  const calls: RecordedCall[] = [];
  const responses: { method: string; value: unknown }[] = [];

  const api = vi.fn((path: string) => {
    const call: RecordedCall = { path, method: "get", headers: {} };

    const builder = {
      header(key: string, value: string) {
        call.headers[key] = value;
        return this;
      },
      query(q: Record<string, unknown>) {
        call.query = q;
        return this;
      },
      select(s: string) {
        call.select = s;
        return this;
      },
      top(n: number) {
        call.top = n;
        return this;
      },
      filter(f: string) {
        call.filter = f;
        return this;
      },
      orderby(o: string) {
        call.orderby = o;
        return this;
      },
      search(s: string) {
        call.search = s;
        return this;
      },
      async get() {
        call.method = "get";
        calls.push(call);
        return popResponse(responses, "get");
      },
      async post(body: unknown) {
        call.method = "post";
        call.body = body;
        calls.push(call);
        return popResponse(responses, "post");
      },
      async patch(body: unknown) {
        call.method = "patch";
        call.body = body;
        calls.push(call);
        return popResponse(responses, "patch");
      },
      async delete() {
        call.method = "delete";
        calls.push(call);
        return popResponse(responses, "delete");
      },
    };
    return builder;
  });

  return { api, calls, responses };
}

function popResponse(
  responses: { method: string; value: unknown }[],
  method: string,
): unknown {
  const idx = responses.findIndex((r) => r.method === method);
  if (idx === -1) {
    throw new Error(
      `mock: no response queued for ${method.toUpperCase()} (next queued: ${
        responses.map((r) => r.method).join(",") || "none"
      })`,
    );
  }
  return responses.splice(idx, 1)[0].value;
}
