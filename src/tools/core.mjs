import { z, tool } from './common.mjs';

export function registerCore(server, graph) {
  tool(server, 'whoami', { title: 'Signed-in account', description: 'Shows the Microsoft 365 account this connection is signed in as.' }, () =>
    graph('GET', '/me', { query: { $select: 'displayName,userPrincipalName,mail,id,jobTitle,officeLocation' } }),
  );

  tool(
    server,
    'graph_request',
    {
      title: 'Call Microsoft Graph',
      description:
        'Calls any Microsoft Graph endpoint as the signed-in user, for anything the other tools do not cover (rules, mailbox settings, groups, presence, online meetings, sharing links, ...), and to fetch the next page of a list via its nextLink. Can change or delete data and send messages, so be careful with non-GET methods.',
      kind: 'destructive',
      input: {
        method: z.enum(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']).default('GET'),
        path: z.string().describe('Path after the version, e.g. /me/mailboxSettings or /users/info@example.com/mailFolders/inbox/messageRules, or a full https://graph.microsoft.com nextLink URL'),
        query: z.record(z.string(), z.string()).optional().describe('Query parameters, e.g. {"$select":"id,subject","$top":"10"}'),
        body: z.any().optional().describe('JSON request body'),
        headers: z.record(z.string(), z.string()).optional().describe('Extra headers, e.g. {"ConsistencyLevel":"eventual"} or {"If-Match":"<etag>"}'),
        api_version: z.enum(['v1.0', 'beta']).default('v1.0'),
      },
    },
    async ({ method, path, query, body, headers, api_version }) => (await graph(method, path, { query, body, headers, version: api_version })) ?? { status: 'done' },
  );
}
