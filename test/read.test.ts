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

describe("read tools", () => {
  it("listFolders maps Graph response into a flat shape", async () => {
    const { listFolders } = await import("../src/tools/read.js");
    mock.responses.push({
      method: "get",
      value: {
        value: [
          {
            id: "f1",
            displayName: "Inbox",
            parentFolderId: "root",
            unreadItemCount: 5,
            totalItemCount: 100,
            childFolderCount: 1,
          },
        ],
      },
    });
    const out = (await listFolders({})) as { folders: unknown[] };
    expect(out.folders).toEqual([
      {
        id: "f1",
        name: "Inbox",
        parentFolderId: "root",
        unread: 5,
        total: 100,
        hasChildren: true,
      },
    ]);
    expect(mock.calls[0].path).toBe("/me/mailFolders");
  });

  it("listRecent applies orderby and unreadOnly filter", async () => {
    const { listRecent } = await import("../src/tools/read.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    await listRecent({ folder: "inbox", limit: 5, unreadOnly: true });
    const c = mock.calls[0];
    expect(c.path).toBe("/me/mailFolders/inbox/messages");
    expect(c.orderby).toBe("receivedDateTime DESC");
    expect(c.filter).toBe("isRead eq false");
    expect(c.top).toBe(5);
  });

  it("listRecent skips filter when unreadOnly is false", async () => {
    const { listRecent } = await import("../src/tools/read.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    await listRecent({ folder: "inbox", limit: 25, unreadOnly: false });
    expect(mock.calls[0].filter).toBeUndefined();
  });

  it("search uses $search and escapes embedded quotes", async () => {
    const { search } = await import("../src/tools/read.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    await search({ query: 'a "b" c', limit: 10 });
    const c = mock.calls[0];
    expect(c.path).toBe("/me/messages");
    expect(c.search).toBe('"a \\"b\\" c"');
    // Graph forbids combining $search with $orderby — make sure we don't add one.
    expect(c.orderby).toBeUndefined();
  });

  it("search escapes backslashes BEFORE quotes (CodeQL js/incomplete-sanitization)", async () => {
    const { search } = await import("../src/tools/read.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    // Input contains a quote AND a literal backslash. If we escaped only quotes,
    // a query like `a\"b` would let the embedded quote escape early.
    await search({ query: 'a\\"b', limit: 1 });
    const c = mock.calls[0];
    // Expected: backslash doubled, then quote escaped: `a\\\"b`.
    expect(c.search).toBe('"a\\\\\\"b"');
  });

  // A realistic Graph @odata.nextLink — absolute URL on the graph host.
  const NEXT_LINK =
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=1";

  it("listRecent returns nextCursor from @odata.nextLink and follows it via cursor", async () => {
    const { listRecent } = await import("../src/tools/read.js");
    mock.responses.push({
      method: "get",
      value: { value: [{ id: "m1" }], "@odata.nextLink": NEXT_LINK },
    });
    const page1 = (await listRecent({
      folder: "inbox",
      limit: 1,
      unreadOnly: false,
    })) as { messages: unknown[]; nextCursor: string | null };
    expect(page1.nextCursor).toBe(NEXT_LINK);

    // Following the cursor hits the opaque URL directly, not the folder path.
    mock.responses.push({
      method: "get",
      value: { value: [{ id: "m2" }] },
    });
    const page2 = (await listRecent({
      folder: "inbox",
      limit: 1,
      unreadOnly: false,
      cursor: NEXT_LINK,
    })) as { messages: unknown[]; nextCursor: string | null };
    expect(mock.calls[1].path).toBe(NEXT_LINK);
    expect(mock.calls[1].top).toBeUndefined(); // cursor encodes everything
    expect(page2.nextCursor).toBeNull();
  });

  it("search returns nextCursor null when no more pages", async () => {
    const { search } = await import("../src/tools/read.js");
    mock.responses.push({ method: "get", value: { value: [] } });
    const out = (await search({ query: "x", limit: 10 })) as {
      nextCursor: string | null;
    };
    expect(out.nextCursor).toBeNull();
  });

  it("rejects a cursor pointing at a non-Graph host (SSRF guard) without hitting the network", async () => {
    const { listRecent, search } = await import("../src/tools/read.js");
    await expect(
      listRecent({
        folder: "inbox",
        limit: 1,
        unreadOnly: false,
        cursor: "https://evil.example.com/steal",
      }),
    ).rejects.toThrow(/not an https Microsoft Graph endpoint/);
    await expect(
      search({ query: "x", limit: 1, cursor: "not a url" }),
    ).rejects.toThrow(/not a URL/);
    // The guard short-circuits before any Graph call is made.
    expect(mock.calls.length).toBe(0);
  });

  it("rejects a host-suffix spoof and a non-https Graph cursor", async () => {
    const { assertGraphCursor } = await import("../src/tools/read.js");
    expect(() =>
      assertGraphCursor("https://graph.microsoft.com.evil.com/x"),
    ).toThrow();
    expect(() => assertGraphCursor("http://graph.microsoft.com/x")).toThrow();
    // Sovereign-cloud + global hosts pass through unchanged.
    expect(assertGraphCursor("https://graph.microsoft.us/v1.0/me")).toBe(
      "https://graph.microsoft.us/v1.0/me",
    );
  });

  it("read truncates an oversized body and flags truncated", async () => {
    const { read } = await import("../src/tools/read.js");
    const big = "x".repeat(50);
    mock.responses.push({
      method: "get",
      value: { id: "m1", body: { contentType: "text", content: big } },
    });
    const out = (await read({
      messageId: "m1",
      bodyFormat: "text",
      maxBodyChars: 10,
    })) as { body: { content: string; truncated: boolean } };
    expect(out.body.truncated).toBe(true);
    expect(out.body.content).toBe("xxxxxxxxxx\n…[truncated]");
  });

  it("read does not split a surrogate pair when truncating", async () => {
    const { read } = await import("../src/tools/read.js");
    // "ab" + 😀 (a surrogate pair). Cap at 3 lands between the pair's halves.
    const body = "ab\u{1F600}cd";
    mock.responses.push({
      method: "get",
      value: { id: "m1", body: { contentType: "text", content: body } },
    });
    const out = (await read({
      messageId: "m1",
      bodyFormat: "text",
      maxBodyChars: 3,
    })) as { body: { content: string; truncated: boolean } };
    expect(out.body.truncated).toBe(true);
    // The dangling high surrogate is dropped, so the prefix is valid UTF-16.
    expect(out.body.content).toBe("ab\n…[truncated]");
    expect(out.body.content).not.toContain("�");
  });

  it("read does not discard prose that follows a large leading style block", async () => {
    // Regression: a pre-trim optimization once cut inside a big leading
    // <style> whose close tag fell past the cut, silently dropping the body.
    const { read } = await import("../src/tools/read.js");
    const html =
      "<style>" + "Z".repeat(5000) + "</style>IMPORTANT PROSE the user wants";
    mock.responses.push({
      method: "get",
      value: { id: "m1", body: { contentType: "html", content: html } },
    });
    const out = (await read({
      messageId: "m1",
      bodyFormat: "text",
      maxBodyChars: 50,
    })) as { body: { content: string } };
    expect(out.body.content).toContain("IMPORTANT PROSE");
  });

  it("read strips HTML to text when text was requested but Graph returned html", async () => {
    const { read } = await import("../src/tools/read.js");
    mock.responses.push({
      method: "get",
      value: {
        id: "m1",
        body: { contentType: "html", content: "<p>Hi <b>there</b></p>" },
      },
    });
    const out = (await read({
      messageId: "m1",
      bodyFormat: "text",
      maxBodyChars: 0,
    })) as { body: { content: string; contentType: string } };
    expect(out.body.contentType).toBe("text");
    expect(out.body.content).toBe("Hi there");
  });

  it("read returns full message with formatted recipients and body", async () => {
    const { read } = await import("../src/tools/read.js");
    mock.responses.push({
      method: "get",
      value: {
        id: "m1",
        subject: "hi",
        from: { emailAddress: { name: "A", address: "a@x.com" } },
        toRecipients: [{ emailAddress: { name: "B", address: "b@x.com" } }],
        ccRecipients: [],
        replyTo: [],
        receivedDateTime: "2026-05-20T00:00:00Z",
        body: { contentType: "text", content: "hello" },
      },
    });
    const out = (await read({ messageId: "m1", bodyFormat: "text" })) as Record<
      string,
      unknown
    >;
    expect(out.from).toBe("A <a@x.com>");
    expect(out.to).toEqual(["B <b@x.com>"]);
    expect((out.body as { content: string }).content).toBe("hello");
    expect(mock.calls[0].headers["Prefer"]).toBe(
      'outlook.body-content-type="text"',
    );
  });
});
