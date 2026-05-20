// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { describe, it, expect, beforeEach, vi } from "vitest";
import { makeMockGraph } from "./_mockGraph.js";

const mock = makeMockGraph();

vi.mock("../src/graph.js", () => ({
  graph: { api: mock.api },
  getMe: vi.fn(),
}));

beforeEach(() => {
  mock.calls.length = 0;
  mock.responses.length = 0;
  mock.api.mockClear();
});

describe("mutate tools", () => {
  it("markRead PATCHes isRead field", async () => {
    const { markRead } = await import("../src/tools/mutate.js");
    mock.responses.push({ method: "patch", value: {} });
    const out = (await markRead({ messageId: "m1", isRead: true })) as Record<
      string,
      unknown
    >;
    expect(mock.calls[0].method).toBe("patch");
    expect(mock.calls[0].path).toBe("/me/messages/m1");
    expect(mock.calls[0].body).toEqual({ isRead: true });
    expect(out.ok).toBe(true);
  });

  it("move POSTs to /move with destinationId", async () => {
    const { move } = await import("../src/tools/mutate.js");
    mock.responses.push({ method: "post", value: { id: "newId" } });
    const out = (await move({
      messageId: "m1",
      destinationFolder: "archive",
    })) as Record<string, unknown>;
    expect(mock.calls[0].path).toBe("/me/messages/m1/move");
    expect(mock.calls[0].body).toEqual({ destinationId: "archive" });
    // Graph returns the message with a NEW id; we surface both.
    expect(out.originalMessageId).toBe("m1");
    expect(out.newMessageId).toBe("newId");
  });

  it("delete in soft mode moves to deleteditems", async () => {
    const { deleteMessage } = await import("../src/tools/mutate.js");
    mock.responses.push({ method: "post", value: { id: "binId" } });
    const out = (await deleteMessage({
      messageId: "m1",
      hardDelete: false,
    })) as Record<string, unknown>;
    expect(mock.calls[0].path).toBe("/me/messages/m1/move");
    expect(mock.calls[0].body).toEqual({ destinationId: "deleteditems" });
    expect(out.mode).toBe("soft");
    expect(out.newMessageId).toBe("binId");
  });

  it("delete in hard mode DELETEs", async () => {
    const { deleteMessage } = await import("../src/tools/mutate.js");
    mock.responses.push({ method: "delete", value: undefined });
    const out = (await deleteMessage({
      messageId: "m1",
      hardDelete: true,
    })) as Record<string, unknown>;
    expect(mock.calls[0].method).toBe("delete");
    expect(mock.calls[0].path).toBe("/me/messages/m1");
    expect(out.mode).toBe("hard");
  });
});
