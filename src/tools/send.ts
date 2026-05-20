// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { z } from "zod";
import { graph } from "../graph.js";

const recipientList = z
  .array(z.string().email())
  .min(1)
  .describe("One or more email addresses.");

const optionalRecipientList = z.array(z.string().email()).default([]);

const bodyFormat = z
  .enum(["text", "html"])
  .default("text")
  .describe(
    "Body content type. 'text' is plain text; 'html' is rendered HTML.",
  );

function toRecipients(
  addrs: string[],
): { emailAddress: { address: string } }[] {
  return addrs.map((address) => ({ emailAddress: { address } }));
}

export const sendSchema = z.object({
  to: recipientList,
  cc: optionalRecipientList,
  bcc: optionalRecipientList,
  subject: z.string().default(""),
  body: z.string().default(""),
  bodyFormat,
});

export type SendInput = z.infer<typeof sendSchema>;

export async function send(input: SendInput): Promise<unknown> {
  // sendMail saves a copy to Sent Items by default (saveToSentItems=true).
  await graph.api("/me/sendMail").post({
    message: {
      subject: input.subject,
      body: { contentType: input.bodyFormat, content: input.body },
      toRecipients: toRecipients(input.to),
      ccRecipients: toRecipients(input.cc),
      bccRecipients: toRecipients(input.bcc),
    },
    saveToSentItems: true,
  });

  return {
    ok: true,
    to: input.to,
    cc: input.cc,
    bcc: input.bcc,
    subject: input.subject,
  };
}

export const replySchema = z.object({
  messageId: z.string().min(1),
  body: z
    .string()
    .min(1)
    .describe("Reply body. Prepended to the original quoted thread by Graph."),
  bodyFormat,
  replyAll: z
    .boolean()
    .default(false)
    .describe("If true, replies to all recipients of the original message."),
});

export type ReplyInput = z.infer<typeof replySchema>;

export async function reply(input: ReplyInput): Promise<unknown> {
  const path = input.replyAll
    ? `/me/messages/${encodeURIComponent(input.messageId)}/replyAll`
    : `/me/messages/${encodeURIComponent(input.messageId)}/reply`;

  await graph.api(path).post({
    message: {
      body: { contentType: input.bodyFormat, content: input.body },
    },
  });

  return {
    ok: true,
    messageId: input.messageId,
    replyAll: input.replyAll,
  };
}

export const createDraftSchema = z.object({
  to: recipientList,
  cc: optionalRecipientList,
  bcc: optionalRecipientList,
  subject: z.string().default(""),
  body: z.string().default(""),
  bodyFormat,
});

export type CreateDraftInput = z.infer<typeof createDraftSchema>;

export async function createDraft(input: CreateDraftInput): Promise<unknown> {
  // POST /me/messages creates a message in the Drafts folder.
  const draft = await graph.api("/me/messages").post({
    subject: input.subject,
    body: { contentType: input.bodyFormat, content: input.body },
    toRecipients: toRecipients(input.to),
    ccRecipients: toRecipients(input.cc),
    bccRecipients: toRecipients(input.bcc),
  });

  return {
    ok: true,
    draftId: draft?.id,
    webLink: draft?.webLink,
    subject: draft?.subject,
  };
}

export const sendDraftSchema = z.object({
  draftId: z
    .string()
    .min(1)
    .describe(
      "Id of an existing draft (returned by personal_email_create_draft).",
    ),
});

export type SendDraftInput = z.infer<typeof sendDraftSchema>;

export async function sendDraft(input: SendDraftInput): Promise<unknown> {
  await graph
    .api(`/me/messages/${encodeURIComponent(input.draftId)}/send`)
    .post({});
  return { ok: true, draftId: input.draftId };
}
