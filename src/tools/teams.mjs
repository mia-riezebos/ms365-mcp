import { z, tool, enc, page, top } from './common.mjs';

const summarizeMessage = (m) => ({
  id: m.id,
  replyToId: m.replyToId || undefined,
  from: m.from?.user?.displayName || m.from?.application?.displayName,
  created: m.createdDateTime,
  subject: m.subject || undefined,
  body: m.body?.contentType === 'html' ? m.body.content.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim() : m.body?.content,
  attachments: m.attachments?.length ? m.attachments.map((a) => a.name || a.contentType) : undefined,
  messageType: m.messageType !== 'message' ? m.messageType : undefined,
  webUrl: m.webUrl,
});
const content = { content: z.string(), content_type: z.enum(['text', 'html']).default('text') };

export function registerTeams(server, graph) {
  tool(server, 'list_teams', { title: 'List my teams', description: 'Lists the Teams the signed-in user is a member of.' }, async () =>
    page(await graph('GET', '/me/joinedTeams', { query: { $select: 'id,displayName,description' } })),
  );

  tool(server, 'list_channels', { title: 'List channels', description: 'Lists the channels of a team.', input: { team_id: z.string() } }, async ({ team_id }) =>
    page(await graph('GET', `/teams/${enc(team_id)}/channels`, { query: { $select: 'id,displayName,description,membershipType,webUrl' } })),
  );

  tool(
    server,
    'list_channel_messages',
    {
      title: 'Read channel messages',
      description: 'Lists recent posts in a channel (newest first), or the replies to one post.',
      input: { team_id: z.string(), channel_id: z.string(), message_id: z.string().optional().describe('Post whose replies to list'), top: top(50, 20) },
    },
    async ({ team_id, channel_id, message_id, top }) => {
      const base = `/teams/${enc(team_id)}/channels/${enc(channel_id)}/messages${message_id ? `/${enc(message_id)}/replies` : ''}`;
      return page(await graph('GET', base, { query: { $top: top } }), summarizeMessage);
    },
  );

  tool(
    server,
    'send_channel_message',
    {
      title: 'Post in a channel',
      description: 'Posts a new message in a channel, or replies to a post.',
      kind: 'destructive',
      input: { team_id: z.string(), channel_id: z.string(), reply_to_message_id: z.string().optional(), subject: z.string().optional(), ...content },
    },
    async ({ team_id, channel_id, reply_to_message_id, subject, content, content_type }) => {
      const base = `/teams/${enc(team_id)}/channels/${enc(channel_id)}/messages${reply_to_message_id ? `/${enc(reply_to_message_id)}/replies` : ''}`;
      return summarizeMessage(await graph('POST', base, { body: { subject, body: { contentType: content_type, content } } }));
    },
  );

  tool(
    server,
    'list_chats',
    { title: 'List chats', description: 'Lists the signed-in user\'s 1:1, group and meeting chats, most recently active first.', input: { top: top(50, 20) } },
    async ({ top }) => {
      const res = await graph('GET', '/me/chats', { query: { $top: top, $orderby: 'lastMessagePreview/createdDateTime desc', $expand: 'members,lastMessagePreview' } });
      return page(res, (c) => ({
        id: c.id,
        topic: c.topic || undefined,
        chatType: c.chatType,
        members: c.members?.map((m) => m.displayName),
        lastMessage: c.lastMessagePreview && { from: c.lastMessagePreview.from?.user?.displayName, at: c.lastMessagePreview.createdDateTime, preview: c.lastMessagePreview.body?.content?.replace(/<[^>]+>/g, '').slice(0, 200) },
        webUrl: c.webUrl,
      }));
    },
  );

  tool(server, 'list_chat_messages', { title: 'Read a chat', description: 'Lists recent messages in a chat, newest first.', input: { chat_id: z.string(), top: top(50, 20) } }, async ({ chat_id, top }) =>
    page(await graph('GET', `/chats/${enc(chat_id)}/messages`, { query: { $top: top, $orderby: 'createdDateTime desc' } }), summarizeMessage),
  );

  tool(server, 'send_chat_message', { title: 'Send a chat message', description: 'Sends a message in a chat.', kind: 'destructive', input: { chat_id: z.string(), ...content } }, async ({ chat_id, content, content_type }) =>
    summarizeMessage(await graph('POST', `/chats/${enc(chat_id)}/messages`, { body: { body: { contentType: content_type, content } } })),
  );
}
