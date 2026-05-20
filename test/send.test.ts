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

describe("send tools", () => {
  it("send posts to /me/sendMail with saveToSentItems and structured recipients", async () => {
    const { send } = await import("../src/tools/send.js");
    mock.responses.push({ method: "post", value: undefined });
    await send({
      to: ["a@x.com"],
      cc: ["b@x.com"],
      bcc: [],
      subject: "hello",
      body: "world",
      bodyFormat: "text",
    });
    const c = mock.calls[0];
    expect(c.path).toBe("/me/sendMail");
    expect(c.body).toEqual({
      message: {
        subject: "hello",
        body: { contentType: "text", content: "world" },
        toRecipients: [{ emailAddress: { address: "a@x.com" } }],
        ccRecipients: [{ emailAddress: { address: "b@x.com" } }],
        bccRecipients: [],
      },
      saveToSentItems: true,
    });
  });

  it("reply chooses /reply or /replyAll based on flag", async () => {
    const { reply } = await import("../src/tools/send.js");
    mock.responses.push({ method: "post", value: undefined });
    await reply({
      messageId: "m1",
      body: "ok",
      bodyFormat: "text",
      replyAll: false,
    });
    expect(mock.calls[0].path).toBe("/me/messages/m1/reply");

    mock.responses.push({ method: "post", value: undefined });
    await reply({
      messageId: "m1",
      body: "ok",
      bodyFormat: "text",
      replyAll: true,
    });
    expect(mock.calls[1].path).toBe("/me/messages/m1/replyAll");
  });

  it("createDraft returns the draft id from POST /me/messages", async () => {
    const { createDraft } = await import("../src/tools/send.js");
    mock.responses.push({
      method: "post",
      value: { id: "draftXYZ", webLink: "https://link", subject: "hi" },
    });
    const out = (await createDraft({
      to: ["a@x.com"],
      cc: [],
      bcc: [],
      subject: "hi",
      body: "",
      bodyFormat: "text",
    })) as Record<string, unknown>;
    expect(mock.calls[0].path).toBe("/me/messages");
    expect(out.draftId).toBe("draftXYZ");
  });

  it("sendDraft POSTs to /send", async () => {
    const { sendDraft } = await import("../src/tools/send.js");
    mock.responses.push({ method: "post", value: undefined });
    await sendDraft({ draftId: "d1" });
    expect(mock.calls[0].path).toBe("/me/messages/d1/send");
  });
});
