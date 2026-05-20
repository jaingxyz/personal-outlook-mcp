// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { z } from "zod";
import { graph } from "../graph.js";

export const markReadSchema = z.object({
  messageId: z.string().min(1),
  isRead: z.boolean().default(true).describe("true to mark read, false to mark unread"),
});

export type MarkReadInput = z.infer<typeof markReadSchema>;

export async function markRead(input: MarkReadInput): Promise<unknown> {
  await graph
    .api(`/me/messages/${encodeURIComponent(input.messageId)}`)
    .patch({ isRead: input.isRead });

  return { ok: true, messageId: input.messageId, isRead: input.isRead };
}

export const moveSchema = z.object({
  messageId: z.string().min(1),
  destinationFolder: z
    .string()
    .min(1)
    .describe(
      "Destination folder id or well-known name: inbox, sentitems, drafts, deleteditems, archive, junkemail.",
    ),
});

export type MoveInput = z.infer<typeof moveSchema>;

export async function move(input: MoveInput): Promise<unknown> {
  // POST /me/messages/{id}/move returns the new message (different id in the destination folder).
  const moved = await graph
    .api(`/me/messages/${encodeURIComponent(input.messageId)}/move`)
    .post({ destinationId: input.destinationFolder });

  return {
    ok: true,
    originalMessageId: input.messageId,
    newMessageId: moved?.id,
    destinationFolder: input.destinationFolder,
  };
}

export const deleteSchema = z.object({
  messageId: z.string().min(1),
  hardDelete: z
    .boolean()
    .default(false)
    .describe(
      "false (default): move to Deleted Items, matching Outlook UI. true: permanently delete — not recoverable.",
    ),
});

export type DeleteInput = z.infer<typeof deleteSchema>;

export async function deleteMessage(input: DeleteInput): Promise<unknown> {
  if (input.hardDelete) {
    await graph
      .api(`/me/messages/${encodeURIComponent(input.messageId)}`)
      .delete();
    return { ok: true, messageId: input.messageId, mode: "hard" };
  }

  const moved = await graph
    .api(`/me/messages/${encodeURIComponent(input.messageId)}/move`)
    .post({ destinationId: "deleteditems" });

  return {
    ok: true,
    originalMessageId: input.messageId,
    newMessageId: moved?.id,
    mode: "soft",
  };
}
