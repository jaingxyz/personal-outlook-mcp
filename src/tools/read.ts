import { z } from "zod";
import { graph } from "../graph.js";

const messageSelect = [
  "id",
  "subject",
  "from",
  "toRecipients",
  "ccRecipients",
  "receivedDateTime",
  "sentDateTime",
  "isRead",
  "hasAttachments",
  "bodyPreview",
  "webLink",
  "conversationId",
  "parentFolderId",
].join(",");

const fullMessageSelect = [
  ...messageSelect.split(","),
  "body",
  "internetMessageHeaders",
  "replyTo",
  "importance",
].join(",");

export const listFoldersSchema = z.object({});

export type ListFoldersInput = z.infer<typeof listFoldersSchema>;

export async function listFolders(
  _input: ListFoldersInput,
): Promise<unknown> {
  const res = await graph
    .api("/me/mailFolders")
    .top(100)
    .select("id,displayName,parentFolderId,unreadItemCount,totalItemCount,childFolderCount")
    .get();

  return {
    folders: (res.value ?? []).map((f: Record<string, unknown>) => ({
      id: f.id,
      name: f.displayName,
      parentFolderId: f.parentFolderId,
      unread: f.unreadItemCount,
      total: f.totalItemCount,
      hasChildren: (f.childFolderCount as number) > 0,
    })),
  };
}

export const listRecentSchema = z.object({
  folder: z
    .string()
    .default("inbox")
    .describe(
      "Folder id or well-known name: inbox, sentitems, drafts, deleteditems, archive, junkemail.",
    ),
  limit: z.number().int().min(1).max(100).default(25),
  unreadOnly: z.boolean().default(false),
});

export type ListRecentInput = z.infer<typeof listRecentSchema>;

export async function listRecent(input: ListRecentInput): Promise<unknown> {
  let req = graph
    .api(`/me/mailFolders/${encodeURIComponent(input.folder)}/messages`)
    .top(input.limit)
    .select(messageSelect)
    .orderby("receivedDateTime DESC");

  if (input.unreadOnly) {
    req = req.filter("isRead eq false");
  }

  const res = await req.get();
  return { messages: (res.value ?? []).map(summarizeMessage) };
}

export const searchSchema = z.object({
  query: z
    .string()
    .min(1)
    .describe(
      "Free-text search across subject, body, and participants. Graph KQL also supported, e.g. 'from:alice@example.com subject:invoice'.",
    ),
  limit: z.number().int().min(1).max(100).default(25),
});

export type SearchInput = z.infer<typeof searchSchema>;

export async function search(input: SearchInput): Promise<unknown> {
  // Graph forbids combining $search with $orderby; results come back ranked by relevance.
  const res = await graph
    .api("/me/messages")
    .top(input.limit)
    .select(messageSelect)
    .search(`"${input.query.replace(/"/g, '\\"')}"`)
    .get();

  return { messages: (res.value ?? []).map(summarizeMessage) };
}

export const readSchema = z.object({
  messageId: z.string().min(1),
  bodyFormat: z.enum(["text", "html"]).default("text"),
});

export type ReadInput = z.infer<typeof readSchema>;

export async function read(input: ReadInput): Promise<unknown> {
  const msg = await graph
    .api(`/me/messages/${encodeURIComponent(input.messageId)}`)
    .header("Prefer", `outlook.body-content-type="${input.bodyFormat}"`)
    .select(fullMessageSelect)
    .get();

  return {
    id: msg.id,
    subject: msg.subject,
    from: formatRecipient(msg.from),
    to: (msg.toRecipients ?? []).map(formatRecipient),
    cc: (msg.ccRecipients ?? []).map(formatRecipient),
    replyTo: (msg.replyTo ?? []).map(formatRecipient),
    receivedDateTime: msg.receivedDateTime,
    sentDateTime: msg.sentDateTime,
    isRead: msg.isRead,
    importance: msg.importance,
    hasAttachments: msg.hasAttachments,
    webLink: msg.webLink,
    conversationId: msg.conversationId,
    parentFolderId: msg.parentFolderId,
    body: {
      contentType: msg.body?.contentType,
      content: msg.body?.content,
    },
  };
}

function summarizeMessage(m: Record<string, unknown>): Record<string, unknown> {
  return {
    id: m.id,
    subject: m.subject,
    from: formatRecipient(m.from as RecipientShape | undefined),
    to: ((m.toRecipients as RecipientShape[]) ?? []).map(formatRecipient),
    receivedDateTime: m.receivedDateTime,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
    preview: m.bodyPreview,
    webLink: m.webLink,
  };
}

type RecipientShape = {
  emailAddress?: { name?: string; address?: string };
};

function formatRecipient(r: RecipientShape | undefined | null): string | null {
  const ea = r?.emailAddress;
  if (!ea) return null;
  if (ea.name && ea.address) return `${ea.name} <${ea.address}>`;
  return ea.address ?? ea.name ?? null;
}
