import { z } from 'zod';
import { mailboxPath } from '../graph.mjs';
import { addr, decodeBase64, mailbox, ok, recipients, toRecipients } from './common.mjs';

const MESSAGE_FIELDS = 'id,subject,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,flag,categories,bodyPreview,conversationId,parentFolderId,webLink';

function summarize(m) {
  return {
    id: m.id,
    subject: m.subject,
    from: addr(m.from),
    sender: m.sender && addr(m.sender) !== addr(m.from) ? addr(m.sender) : undefined,
    to: m.toRecipients?.map(addr),
    cc: m.ccRecipients?.length ? m.ccRecipients.map(addr) : undefined,
    received: m.receivedDateTime,
    isRead: m.isRead,
    isDraft: m.isDraft || undefined,
    hasAttachments: m.hasAttachments,
    importance: m.importance !== 'normal' ? m.importance : undefined,
    flag: m.flag?.flagStatus !== 'notFlagged' ? m.flag?.flagStatus : undefined,
    categories: m.categories?.length ? m.categories : undefined,
    preview: m.bodyPreview,
    conversationId: m.conversationId,
    webLink: m.webLink,
  };
}

function wrap(handler) {
  return async (args) => {
    try {
      return await handler(args);
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: err.message }] };
    }
  };
}

const folderPath = (mb, folder) => `${mailboxPath(mb)}/mailFolders/${encodeURIComponent(folder)}`;

export function registerMail(server, graph) {
  const read = { readOnlyHint: true, openWorldHint: true };
  const write = { readOnlyHint: false, destructiveHint: false, openWorldHint: true };

  server.registerTool(
    'list_mail_folders',
    {
      title: 'List mail folders',
      description: 'Lists mail folders in a mailbox (top level, or the children of parent_folder), with unread and total counts.',
      inputSchema: {
        mailbox,
        parent_folder: z.string().optional().describe('Folder id or well-known name (inbox, sentitems, drafts, deleteditems, archive, junkemail) whose child folders to list'),
        top: z.number().int().min(1).max(250).default(100),
      },
      annotations: read,
    },
    wrap(async ({ mailbox: mb, parent_folder, top }) => {
      const path = parent_folder ? `${folderPath(mb, parent_folder)}/childFolders` : `${mailboxPath(mb)}/mailFolders`;
      const res = await graph('GET', path, {
        query: { $top: top, $select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount' },
      });
      return ok(res.value);
    }),
  );

  server.registerTool(
    'list_messages',
    {
      title: 'List or search messages',
      description:
        'Lists messages in a mailbox folder, newest first, or searches them. Works on the signed-in mailbox and on shared mailboxes the account has Full Access to.',
      inputSchema: {
        mailbox,
        folder: z.string().optional().describe('Folder id or well-known name (inbox, sentitems, drafts, deleteditems, archive, junkemail). Omit to cover all folders.'),
        search: z.string().optional().describe('KQL search, e.g. "invoice", "from:alice@example.com", "subject:report received>=2026-01-01". Results are by relevance, not date.'),
        filter: z.string().optional().describe('OData $filter, e.g. "isRead eq false" or "receivedDateTime ge 2026-01-01T00:00:00Z". Ignored when search is set.'),
        unread_only: z.boolean().default(false),
        top: z.number().int().min(1).max(50).default(10),
        skip: z.number().int().min(0).default(0).describe('Messages to skip, for paging (not with search)'),
      },
      annotations: read,
    },
    wrap(async ({ mailbox: mb, folder, search, filter, unread_only, top, skip }) => {
      const path = folder ? `${folderPath(mb, folder)}/messages` : `${mailboxPath(mb)}/messages`;
      const query = { $top: top, $select: MESSAGE_FIELDS };
      if (search) {
        query.$search = `"${search.replace(/"/g, '\\"')}"`;
      } else {
        const filters = [filter, unread_only ? 'isRead eq false' : undefined].filter(Boolean);
        if (filters.length) query.$filter = filters.map((f) => `(${f})`).join(' and ');
        if (!filter) query.$orderby = 'receivedDateTime desc';
        if (skip) query.$skip = skip;
      }
      const res = await graph('GET', path, { query });
      return ok({ mailbox: mb, folder: folder || 'all', count: res.value.length, hasMore: Boolean(res['@odata.nextLink']), messages: res.value.map(summarize) });
    }),
  );

  server.registerTool(
    'get_message',
    {
      title: 'Read a message',
      description: 'Gets one message with its full body.',
      inputSchema: {
        mailbox,
        message_id: z.string(),
        body_format: z.enum(['text', 'html']).default('text'),
      },
      annotations: read,
    },
    wrap(async ({ mailbox: mb, message_id, body_format }) => {
      const m = await graph('GET', `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}`, {
        query: { $select: `${MESSAGE_FIELDS},bccRecipients,replyTo,body,internetMessageId` },
        headers: { Prefer: `outlook.body-content-type="${body_format}"` },
      });
      return ok({ ...summarize(m), bcc: m.bccRecipients?.map(addr), replyTo: m.replyTo?.map(addr), internetMessageId: m.internetMessageId, body: m.body?.content, preview: undefined });
    }),
  );

  server.registerTool(
    'list_attachments',
    {
      title: 'List attachments',
      description: 'Lists the attachments of a message (metadata only).',
      inputSchema: { mailbox, message_id: z.string() },
      annotations: read,
    },
    wrap(async ({ mailbox: mb, message_id }) => {
      const res = await graph('GET', `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}/attachments`, {
        query: { $select: 'id,name,contentType,size,isInline,lastModifiedDateTime' },
      });
      return ok(res.value);
    }),
  );

  server.registerTool(
    'get_attachment',
    {
      title: 'Download an attachment',
      description: 'Downloads a file attachment (up to 3 MB) as base64, or decoded text for text/* types.',
      inputSchema: { mailbox, message_id: z.string(), attachment_id: z.string() },
      annotations: read,
    },
    wrap(async ({ mailbox: mb, message_id, attachment_id }) => {
      const a = await graph('GET', `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}/attachments/${encodeURIComponent(attachment_id)}`);
      if (!a.contentBytes) return ok({ name: a.name, contentType: a.contentType, size: a.size, note: 'Not a file attachment (item or reference attachment); no bytes returned.' });
      if (a.size > 3 * 1024 * 1024) return ok({ name: a.name, size: a.size, note: 'Attachment larger than 3 MB; not returned.' });
      const isText = a.contentType?.startsWith('text/') || /json|xml|csv/.test(a.contentType || '');
      return ok({ name: a.name, contentType: a.contentType, size: a.size, ...(isText ? { text: new TextDecoder().decode(decodeBase64(a.contentBytes)) } : { contentBase64: a.contentBytes }) });
    }),
  );

  const composeInput = {
    mailbox: mailbox.describe('Mailbox to send from: "me", or a shared mailbox address the account has Send As rights on'),
    to: recipients,
    cc: recipients.optional(),
    bcc: recipients.optional(),
    subject: z.string(),
    body: z.string(),
    body_type: z.enum(['text', 'html']).default('text'),
  };
  const composeMessage = ({ to, cc, bcc, subject, body, body_type }) => ({
    subject,
    body: { contentType: body_type === 'html' ? 'HTML' : 'Text', content: body },
    toRecipients: toRecipients(to),
    ccRecipients: toRecipients(cc),
    bccRecipients: toRecipients(bcc),
  });

  server.registerTool(
    'create_draft',
    {
      title: 'Create a draft',
      description: 'Saves a new draft in the Drafts folder of a mailbox without sending it.',
      inputSchema: composeInput,
      annotations: write,
    },
    wrap(async (args) => {
      const m = await graph('POST', `${mailboxPath(args.mailbox)}/messages`, { body: composeMessage(args) });
      return ok({ created: true, mailbox: args.mailbox, ...summarize(m) });
    }),
  );

  server.registerTool(
    'send_mail',
    {
      title: 'Send an email',
      description: 'Sends an email from a mailbox. Sending from a shared mailbox needs Send As rights on it; the copy is saved in that mailbox\'s Sent Items.',
      inputSchema: { ...composeInput, save_to_sent_items: z.boolean().default(true) },
      annotations: { ...write, destructiveHint: true },
    },
    wrap(async (args) => {
      await graph('POST', `${mailboxPath(args.mailbox)}/sendMail`, {
        body: { message: composeMessage(args), saveToSentItems: args.save_to_sent_items },
      });
      return ok({ sent: true, from: args.mailbox, to: args.to, subject: args.subject });
    }),
  );

  server.registerTool(
    'send_draft',
    {
      title: 'Send a draft',
      description: 'Sends an existing draft message.',
      inputSchema: { mailbox, message_id: z.string() },
      annotations: { ...write, destructiveHint: true },
    },
    wrap(async ({ mailbox: mb, message_id }) => {
      await graph('POST', `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}/send`);
      return ok({ sent: true, mailbox: mb, message_id });
    }),
  );

  server.registerTool(
    'reply_message',
    {
      title: 'Reply to a message',
      description: 'Replies (or replies all) to a message. By default creates a reply draft; set send=true to send immediately.',
      inputSchema: {
        mailbox,
        message_id: z.string(),
        comment: z.string().describe('Reply text placed above the quoted message'),
        reply_all: z.boolean().default(false),
        send: z.boolean().default(false),
      },
      annotations: { ...write, destructiveHint: true },
    },
    wrap(async ({ mailbox: mb, message_id, comment, reply_all, send }) => {
      const base = `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}`;
      const action = reply_all ? 'replyAll' : 'reply';
      if (send) {
        await graph('POST', `${base}/${action}`, { body: { comment } });
        return ok({ sent: true, mailbox: mb, replied_to: message_id });
      }
      const draft = await graph('POST', `${base}/${reply_all ? 'createReplyAll' : 'createReply'}`, { body: { comment } });
      return ok({ draftCreated: true, mailbox: mb, ...summarize(draft) });
    }),
  );

  server.registerTool(
    'forward_message',
    {
      title: 'Forward a message',
      description: 'Forwards a message. By default creates a forward draft; set send=true to send immediately.',
      inputSchema: { mailbox, message_id: z.string(), to: recipients, comment: z.string().default(''), send: z.boolean().default(false) },
      annotations: { ...write, destructiveHint: true },
    },
    wrap(async ({ mailbox: mb, message_id, to, comment, send }) => {
      const base = `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}`;
      const body = { comment, toRecipients: toRecipients(to) };
      if (send) {
        await graph('POST', `${base}/forward`, { body });
        return ok({ sent: true, mailbox: mb, forwarded: message_id, to });
      }
      const draft = await graph('POST', `${base}/createForward`, { body });
      return ok({ draftCreated: true, mailbox: mb, ...summarize(draft) });
    }),
  );

  server.registerTool(
    'move_message',
    {
      title: 'Move a message',
      description: 'Moves a message to another folder in the same mailbox (use deleteditems to delete it, archive to archive it).',
      inputSchema: { mailbox, message_id: z.string(), destination: z.string().describe('Folder id or well-known name (inbox, archive, deleteditems, junkemail, ...)') },
      annotations: write,
    },
    wrap(async ({ mailbox: mb, message_id, destination }) => {
      const m = await graph('POST', `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}/move`, { body: { destinationId: destination } });
      return ok({ moved: true, mailbox: mb, destination, new_message_id: m.id });
    }),
  );

  server.registerTool(
    'update_message',
    {
      title: 'Mark, flag or categorize a message',
      description: 'Updates read state, follow-up flag, importance or categories of a message.',
      inputSchema: {
        mailbox,
        message_id: z.string(),
        is_read: z.boolean().optional(),
        flag: z.enum(['notFlagged', 'flagged', 'complete']).optional(),
        importance: z.enum(['low', 'normal', 'high']).optional(),
        categories: z.array(z.string()).optional().describe('Replaces the message\'s categories'),
      },
      annotations: write,
    },
    wrap(async ({ mailbox: mb, message_id, is_read, flag, importance, categories }) => {
      const body = {};
      if (is_read !== undefined) body.isRead = is_read;
      if (flag) body.flag = { flagStatus: flag };
      if (importance) body.importance = importance;
      if (categories) body.categories = categories;
      const m = await graph('PATCH', `${mailboxPath(mb)}/messages/${encodeURIComponent(message_id)}`, { body });
      return ok({ updated: true, mailbox: mb, ...summarize(m) });
    }),
  );

}
