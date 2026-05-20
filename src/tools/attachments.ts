// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
import { z } from "zod";
import { mkdir, writeFile } from "node:fs/promises";
import { resolve, join } from "node:path";
import { homedir } from "node:os";
import { graph } from "../graph.js";

const DEFAULT_DOWNLOAD_DIR = join(
  homedir(),
  "Downloads",
  "personal-outlook-mcp",
);

export const listAttachmentsSchema = z.object({
  messageId: z.string().min(1),
});

export type ListAttachmentsInput = z.infer<typeof listAttachmentsSchema>;

export async function listAttachments(
  input: ListAttachmentsInput,
): Promise<unknown> {
  const res = await graph
    .api(`/me/messages/${encodeURIComponent(input.messageId)}/attachments`)
    .select("id,name,contentType,size,isInline")
    .get();

  return {
    messageId: input.messageId,
    attachments: (res.value ?? []).map((a: Record<string, unknown>) => ({
      id: a.id,
      name: a.name,
      contentType: a.contentType,
      size: a.size,
      isInline: a.isInline,
      kind: a["@odata.type"],
    })),
  };
}

export const downloadAttachmentSchema = z.object({
  messageId: z.string().min(1),
  attachmentId: z.string().min(1),
  destDir: z
    .string()
    .optional()
    .describe(
      `Directory to save into. Defaults to ${DEFAULT_DOWNLOAD_DIR}. Created if missing. ~ is expanded.`,
    ),
  filename: z
    .string()
    .optional()
    .describe(
      "Override the filename. Defaults to the attachment's name as reported by Graph.",
    ),
});

export type DownloadAttachmentInput = z.infer<typeof downloadAttachmentSchema>;

export async function downloadAttachment(
  input: DownloadAttachmentInput,
): Promise<unknown> {
  const att = await graph
    .api(
      `/me/messages/${encodeURIComponent(input.messageId)}/attachments/${encodeURIComponent(input.attachmentId)}`,
    )
    .get();

  const odataType = att["@odata.type"] as string | undefined;
  if (odataType !== "#microsoft.graph.fileAttachment") {
    throw new Error(
      `Unsupported attachment type: ${odataType ?? "unknown"}. Only fileAttachment is supported (item and reference attachments are not).`,
    );
  }

  const contentBytes = att.contentBytes as string | undefined;
  if (!contentBytes) {
    throw new Error("Attachment has no contentBytes payload.");
  }

  const buffer = Buffer.from(contentBytes, "base64");
  const safeName = sanitizeFilename(
    input.filename ?? (att.name as string) ?? "attachment",
  );
  const dir = expandHome(input.destDir ?? DEFAULT_DOWNLOAD_DIR);
  await mkdir(dir, { recursive: true });
  const fullPath = resolve(dir, safeName);
  await writeFile(fullPath, buffer);

  return {
    ok: true,
    path: fullPath,
    size: buffer.length,
    name: safeName,
    contentType: att.contentType,
  };
}

export function sanitizeFilename(name: string): string {
  // Strip path separators and control chars; the control-char range is the
  // point of this regex.
  // eslint-disable-next-line no-control-regex
  return name.replace(/[/\\\x00-\x1f]+/g, "_").trim() || "attachment";
}

export function expandHome(p: string): string {
  if (p === "~") return homedir();
  if (p.startsWith("~/")) return join(homedir(), p.slice(2));
  return p;
}
