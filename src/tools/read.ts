// SPDX-License-Identifier: AGPL-3.0-or-later
// Copyright (C) 2026 jaingxyz
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

// A pagination `cursor` is an opaque @odata.nextLink we hand straight to
// graph.api(). The Graph SDK treats an absolute https:// URL as the request
// host, so an unvalidated cursor is an SSRF / token-leak vector: a non-Graph
// host would still be fetched (token stripped), and a crafted graph.microsoft
// URL would receive the user's bearer token with attacker-chosen query params.
// Only accept nextLinks pointing at a known Microsoft Graph host (global +
// sovereign clouds). Reject anything else before it reaches the network.
const GRAPH_HOSTS = new Set([
  "graph.microsoft.com",
  "graph.microsoft.us",
  "dod-graph.microsoft.us",
  "graph.microsoft.de",
  "microsoftgraph.chinacloudapi.cn",
]);

export function assertGraphCursor(cursor: string): string {
  let url: URL;
  try {
    url = new URL(cursor);
  } catch {
    throw new Error(
      "Invalid pagination cursor: not a URL. Pass back the exact nextCursor from a previous call.",
    );
  }
  if (url.protocol !== "https:" || !GRAPH_HOSTS.has(url.host.toLowerCase())) {
    throw new Error(
      `Invalid pagination cursor: "${url.protocol}//${url.host}" is not an https Microsoft Graph endpoint. Pass back the exact nextCursor from a previous call.`,
    );
  }
  return cursor;
}

export const listFoldersSchema = z.object({});

export type ListFoldersInput = z.infer<typeof listFoldersSchema>;

export async function listFolders(_input: ListFoldersInput): Promise<unknown> {
  const res = await graph
    .api("/me/mailFolders")
    .top(100)
    .select(
      "id,displayName,parentFolderId,unreadItemCount,totalItemCount,childFolderCount",
    )
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
  cursor: z
    .string()
    .optional()
    .describe(
      "Opaque pagination cursor from a previous call's nextCursor. When set, returns the next page and ignores folder/limit/unreadOnly.",
    ),
});

export type ListRecentInput = z.infer<typeof listRecentSchema>;

export async function listRecent(input: ListRecentInput): Promise<unknown> {
  if (input.cursor) {
    const res = await graph.api(assertGraphCursor(input.cursor)).get();
    return paged(res);
  }

  let req = graph
    .api(`/me/mailFolders/${encodeURIComponent(input.folder)}/messages`)
    .top(input.limit)
    .select(messageSelect)
    .orderby("receivedDateTime DESC");

  if (input.unreadOnly) {
    req = req.filter("isRead eq false");
  }

  const res = await req.get();
  return paged(res);
}

export const searchSchema = z.object({
  query: z
    .string()
    .min(1)
    .describe(
      "Free-text search across subject, body, and participants. Graph KQL also supported, e.g. 'from:alice@example.com subject:invoice'.",
    ),
  limit: z.number().int().min(1).max(100).default(25),
  cursor: z
    .string()
    .optional()
    .describe(
      "Opaque pagination cursor from a previous call's nextCursor. When set, returns the next page and ignores query/limit.",
    ),
});

export type SearchInput = z.infer<typeof searchSchema>;

export async function search(input: SearchInput): Promise<unknown> {
  if (input.cursor) {
    const res = await graph.api(assertGraphCursor(input.cursor)).get();
    return paged(res);
  }

  // Graph forbids combining $search with $orderby; results come back ranked by relevance.
  // Escape backslashes BEFORE quotes — if we only escape quotes, an input like
  // `a\"b` would let the embedded quote terminate the literal early.
  const escaped = input.query.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  const res = await graph
    .api("/me/messages")
    .top(input.limit)
    .select(messageSelect)
    .search(`"${escaped}"`)
    .get();

  return paged(res);
}

// Default ceiling on returned body length. Marketing emails routinely carry
// 100KB+ of HTML; dumping that verbatim is a token/context sink the model
// rarely needs. Callers can raise it (up to the max) or pass 0 for no cap.
const DEFAULT_MAX_BODY_CHARS = 20000;

export const readSchema = z.object({
  messageId: z.string().min(1),
  bodyFormat: z
    .enum(["text", "html"])
    .default("text")
    .describe(
      "Requested body format. 'text' also strips any HTML Graph returns down to plain text.",
    ),
  maxBodyChars: z
    .number()
    .int()
    .min(0)
    .max(1000000)
    .default(DEFAULT_MAX_BODY_CHARS)
    .describe(
      "Truncate the body to this many characters (a marker is appended when truncated). 0 disables the cap. Defaults to 20000.",
    ),
});

export type ReadInput = z.infer<typeof readSchema>;

export async function read(input: ReadInput): Promise<unknown> {
  const msg = await graph
    .api(`/me/messages/${encodeURIComponent(input.messageId)}`)
    .header("Prefer", `outlook.body-content-type="${input.bodyFormat}"`)
    .select(fullMessageSelect)
    .get();

  let contentType: string | undefined = msg.body?.contentType;
  let content: string = msg.body?.content ?? "";

  // `?? DEFAULT` guards callers that bypass Zod parsing (e.g. tests) so the
  // cap is enforced consistently regardless of how `read` is invoked.
  const max = input.maxBodyChars ?? DEFAULT_MAX_BODY_CHARS;

  // Graph sometimes returns HTML even when we asked for text. If the caller
  // wanted text, strip it down so the model gets readable prose, not markup.
  // stripHtml is linear (it drops <script>/<style> *contents* entirely, so a
  // pre-trim would be unsafe — it could cut inside a leading block and discard
  // the prose after it), and fast enough to run on the full body.
  if (input.bodyFormat === "text" && contentType === "html") {
    content = htmlToText(content);
    contentType = "text";
  }

  const { text: bodyContent, truncated } = truncate(content, max);

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
      contentType,
      content: bodyContent,
      truncated,
    },
  };
}

// Tags that imply a line break when text is extracted, so adjacent blocks
// don't run together. Both opening and closing forms emit a newline.
const BLOCK_TAGS = new Set([
  "p",
  "div",
  "li",
  "tr",
  "ul",
  "ol",
  "table",
  "blockquote",
  "section",
  "article",
  "header",
  "footer",
  "hr",
  "br",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
]);

// Strip HTML structure in a SINGLE LINEAR PASS: remove comments, drop the
// *contents* of <script>/<style> blocks, strip all other tags, and emit a
// newline at block boundaries. Walking the string once with indexOf (no regex
// with unbounded backtracking) keeps this O(n). The previous regex-chain
// version was O(n^2) — adversarial bodies with many unterminated "<" or
// "<script" could pin the single-threaded server's event loop for seconds
// (a 480KB body took ~58s in testing). Don't reintroduce `<[^>]+>`-style
// passes here: under the global flag they rescan to end-of-string per "<".
function stripHtml(html: string): string {
  const lower = html.toLowerCase();
  const n = html.length;
  let out = "";
  let i = 0;
  while (i < n) {
    const lt = html.indexOf("<", i);
    if (lt === -1) {
      out += html.slice(i);
      break;
    }
    out += html.slice(i, lt);

    // HTML comment: skip to the closing "-->" (looking for ">" alone would
    // stop early on a "<!-- a > b -->"-style comment and leak its tail).
    if (lower.startsWith("<!--", lt)) {
      const end = html.indexOf("-->", lt + 4);
      i = end === -1 ? n : end + 3;
      continue;
    }

    const gt = html.indexOf(">", lt + 1);
    if (gt === -1) {
      // No tag terminator left — keep the remainder as literal text so
      // "a < b" (a stray "<") survives instead of being eaten.
      out += html.slice(lt);
      break;
    }

    // Parse "</?  name" from the front of the tag (bounded slice; the name is
    // always near the start, so we never scan the whole tag).
    const nameMatch = /^<\s*(\/?)\s*([a-zA-Z][a-zA-Z0-9]*)/.exec(
      html.slice(lt, Math.min(gt + 1, lt + 32)),
    );
    const isClosing = nameMatch ? nameMatch[1] === "/" : false;
    const name = nameMatch ? nameMatch[2].toLowerCase() : "";

    if (!isClosing && (name === "script" || name === "style")) {
      // Skip everything up to the matching close tag — the content is code,
      // not prose, and must not leak into the text.
      const closeIdx = lower.indexOf(`</${name}`, gt + 1);
      if (closeIdx === -1) {
        i = n;
      } else {
        const closeGt = html.indexOf(">", closeIdx);
        i = closeGt === -1 ? n : closeGt + 1;
      }
      continue;
    }

    if (BLOCK_TAGS.has(name)) out += "\n";
    i = gt + 1;
  }
  return out;
}

// Cheap HTML→text for readability (NOT sanitization — the result is returned
// as text to the model, never rendered). Structural stripping is done by the
// linear stripHtml pass; here we only decode the handful of entities that
// actually show up and collapse whitespace. These remaining regexes are all
// fixed/linear, so they add no backtracking risk.
export function htmlToText(html: string): string {
  return stripHtml(html)
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function truncate(
  text: string,
  max: number,
): { text: string; truncated: boolean } {
  if (max === 0 || text.length <= max) return { text, truncated: false };
  let cut = text.slice(0, max);
  // Don't slice through a surrogate pair: if the cut lands between the high
  // and low halves of an astral codepoint (emoji, some CJK), drop the dangling
  // high surrogate so the result is valid UTF-16, not a U+FFFD replacement.
  const lastCode = cut.charCodeAt(cut.length - 1);
  if (lastCode >= 0xd800 && lastCode <= 0xdbff) {
    cut = cut.slice(0, -1);
  }
  return { text: cut + "\n…[truncated]", truncated: true };
}

// Shape a message-list response: summarized messages plus an opaque cursor
// for the next page (Graph's @odata.nextLink). nextCursor is null when there
// are no more results — the model can stop paging.
function paged(res: { value?: unknown[]; "@odata.nextLink"?: string }): {
  messages: Record<string, unknown>[];
  nextCursor: string | null;
} {
  return {
    messages: (res.value ?? []).map((m) =>
      summarizeMessage(m as Record<string, unknown>),
    ),
    nextCursor: res["@odata.nextLink"] ?? null,
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

export function formatRecipient(
  r: RecipientShape | undefined | null,
): string | null {
  const ea = r?.emailAddress;
  if (!ea) return null;
  if (ea.name && ea.address) return `${ea.name} <${ea.address}>`;
  return ea.address ?? ea.name ?? null;
}
