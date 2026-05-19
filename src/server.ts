import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import {
  listFolders,
  listFoldersSchema,
  listRecent,
  listRecentSchema,
  read,
  readSchema,
  search,
  searchSchema,
} from "./tools/read.js";
import {
  deleteMessage,
  deleteSchema,
  markRead,
  markReadSchema,
  move,
  moveSchema,
} from "./tools/mutate.js";
import {
  createDraft,
  createDraftSchema,
  reply,
  replySchema,
  send,
  sendSchema,
  sendDraft,
  sendDraftSchema,
} from "./tools/send.js";
import {
  downloadAttachment,
  downloadAttachmentSchema,
  listAttachments,
  listAttachmentsSchema,
} from "./tools/attachments.js";

export function buildServer(): McpServer {
  const server = new McpServer({
    name: "personal-outlook-mcp",
    version: "0.1.0",
  });

  server.registerTool(
    "personal_email_list_folders",
    {
      description:
        "List the user's mail folders (Inbox, Sent Items, custom folders, etc.) with id, display name, and unread/total counts.",
      inputSchema: listFoldersSchema.shape,
    },
    async (args) => toolResult(await listFolders(args)),
  );

  server.registerTool(
    "personal_email_list_recent",
    {
      description:
        "List the most recent messages in a folder, newest first. Defaults to inbox.",
      inputSchema: listRecentSchema.shape,
    },
    async (args) => toolResult(await listRecent(args)),
  );

  server.registerTool(
    "personal_email_search",
    {
      description:
        "Search messages across the mailbox by free-text query. Supports Graph KQL operators (from:, subject:, hasAttachment:true, etc.). Results are ranked by relevance, not date.",
      inputSchema: searchSchema.shape,
    },
    async (args) => toolResult(await search(args)),
  );

  server.registerTool(
    "personal_email_read",
    {
      description:
        "Fetch a single message by id, including its body. Use the 'id' field returned by list_recent or search.",
      inputSchema: readSchema.shape,
    },
    async (args) => toolResult(await read(args)),
  );

  server.registerTool(
    "personal_email_mark_read",
    {
      description:
        "Mark a message as read or unread. Pass isRead=false to mark unread.",
      inputSchema: markReadSchema.shape,
      annotations: { destructiveHint: false, idempotentHint: true },
    },
    async (args) => toolResult(await markRead(args)),
  );

  server.registerTool(
    "personal_email_move",
    {
      description:
        "Move a message to another folder (by id or well-known name). The message gets a new id in the destination folder.",
      inputSchema: moveSchema.shape,
      annotations: { destructiveHint: false },
    },
    async (args) => toolResult(await move(args)),
  );

  server.registerTool(
    "personal_email_delete",
    {
      description:
        "Delete a message. By default soft-deletes (moves to Deleted Items). Pass hardDelete=true to permanently delete — not recoverable.",
      inputSchema: deleteSchema.shape,
      annotations: { destructiveHint: true },
    },
    async (args) => toolResult(await deleteMessage(args)),
  );

  server.registerTool(
    "personal_email_send",
    {
      description:
        "Send a new email immediately. A copy is saved to Sent Items. Use create_draft instead if you want to review in Outlook before sending.",
      inputSchema: sendSchema.shape,
      annotations: { destructiveHint: true },
    },
    async (args) => toolResult(await send(args)),
  );

  server.registerTool(
    "personal_email_reply",
    {
      description:
        "Reply to an existing message. Sends immediately. Pass replyAll=true to reply to all recipients of the original.",
      inputSchema: replySchema.shape,
      annotations: { destructiveHint: true },
    },
    async (args) => toolResult(await reply(args)),
  );

  server.registerTool(
    "personal_email_create_draft",
    {
      description:
        "Create a draft email in the Drafts folder without sending. Returns a draftId you can pass to personal_email_send_draft.",
      inputSchema: createDraftSchema.shape,
    },
    async (args) => toolResult(await createDraft(args)),
  );

  server.registerTool(
    "personal_email_send_draft",
    {
      description: "Send a previously created draft by id.",
      inputSchema: sendDraftSchema.shape,
      annotations: { destructiveHint: true },
    },
    async (args) => toolResult(await sendDraft(args)),
  );

  server.registerTool(
    "personal_email_list_attachments",
    {
      description:
        "List attachments on a message, with id, name, contentType, and size in bytes.",
      inputSchema: listAttachmentsSchema.shape,
    },
    async (args) => toolResult(await listAttachments(args)),
  );

  server.registerTool(
    "personal_email_download_attachment",
    {
      description:
        "Download a file attachment to disk. Defaults to ~/Downloads/personal-outlook-mcp/. Returns the absolute path. Item and reference attachments are not supported.",
      inputSchema: downloadAttachmentSchema.shape,
    },
    async (args) => toolResult(await downloadAttachment(args)),
  );

  return server;
}

function toolResult(payload: unknown): {
  content: { type: "text"; text: string }[];
} {
  return {
    content: [{ type: "text", text: JSON.stringify(payload, null, 2) }],
  };
}
