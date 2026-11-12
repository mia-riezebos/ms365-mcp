import { z, tool } from './common.mjs';
import { SCOPE_GUIDE, graphCallRefusal, hasAny, isGraphScope, normalizeScope, resourceOf, union } from '../scopes.mjs';

// conn: { granted, upn, requestScopes(scopes, reason) => url, apiFetch(resource, method, url, init) => Response }
export function registerCore(server, graph, conn) {
  tool(
    server,
    'whoami',
    { title: 'Signed-in account', description: 'Shows the Microsoft 365 account this connection is signed in as, and the permission scopes it has been granted.' },
    async () => ({
      account: await graph('GET', '/me', { query: { $select: 'displayName,userPrincipalName,mail,id,jobTitle,officeLocation' } }),
      scopes: conn.granted,
    }),
  );

  tool(
    server,
    'request_scopes',
    {
      title: 'Request more Microsoft 365 access',
      description: `Asks the user to grant this connection more Microsoft permission scopes. Returns a link for the user to open and sign in with Microsoft; once they have, the tools those scopes enable appear (re-list tools, or reconnect if your client does not refresh them). Scopes already granted are kept. Ask only for what the task needs.\n\n${SCOPE_GUIDE}`,
      kind: 'write',
      input: {
        scopes: z.array(z.string()).min(1).describe('Scopes to add, e.g. ["Mail.Read.Shared", "Calendars.ReadWrite"]'),
        reason: z.string().max(300).optional().describe('One line shown to the user on why the access is needed'),
      },
    },
    async ({ scopes, reason }) => {
      const parsed = scopes.map((s) => [s, normalizeScope(s)]);
      const bad = parsed.filter(([, n]) => !n).map(([s]) => s);
      if (bad.length) throw new Error(`Not a scope: ${bad.join(', ')}`);
      const wanted = union(parsed.map(([, n]) => n)).filter((s) => !['openid', 'profile', 'email', 'offline_access'].includes(s.toLowerCase()));
      const missing = wanted.filter((s) => !hasAny(conn.granted, [s]));
      if (!missing.length) return { status: 'already_granted', scopes: conn.granted };
      const resources = new Set(missing.map(resourceOf));
      return {
        status: 'needs_user',
        sign_in_url: await conn.requestScopes(missing, reason),
        requested: missing,
        message: `Ask the user (${conn.upn}) to open sign_in_url and approve with Microsoft; the link is valid for an hour. Then call whoami to check the granted scopes.${resources.size > 1 ? ' Scopes for APIs other than Graph are checked after sign-in.' : ''}`,
      };
    },
  );

  tool(
    server,
    'graph_request',
    {
      title: 'Call Microsoft Graph',
      description:
        'Calls any Microsoft Graph endpoint as the signed-in user, for anything the other tools do not cover (rules, mailbox settings, groups, presence, online meetings, sharing links, ...), and to fetch the next page of a list via its nextLink. Calls into mail, calendar, contacts, files, Teams, Planner, To Do and OneNote need a matching scope on this connection (see request_scopes). Can change or delete data and send messages, so be careful with non-GET methods.',
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
    async ({ method, path, query, body, headers, api_version }) => {
      const refusal = graphCallRefusal(method, path, conn.granted);
      if (refusal) throw new Error(refusal);
      return (await graph(method, path, { query, body, headers, version: api_version })) ?? { status: 'done' };
    },
  );

  const apis = [...new Set(conn.granted.filter((s) => !isGraphScope(s)).map(resourceOf))].filter((r) => r.startsWith('https://'));
  if (apis.length)
    tool(
      server,
      'api_request',
      {
        title: 'Call another Microsoft API',
        description: `Calls a Microsoft API other than Graph as the signed-in user, with a token for the scopes granted on it. APIs available on this connection: ${apis.join(', ')}. Returns status, headers and body text.`,
        kind: 'destructive',
        input: {
          method: z.enum(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']).default('GET'),
          url: z.string().url().describe(`Full URL on one of: ${apis.join(', ')}`),
          body: z.string().optional().describe('Request body (JSON, XML/SOAP, ...)'),
          content_type: z.string().optional().describe('e.g. application/json or text/xml'),
          headers: z.record(z.string(), z.string()).optional(),
        },
      },
      async ({ method, url, body, content_type, headers }) => {
        const resource = apis.find((r) => new URL(url).origin === new URL(r).origin);
        if (!resource) throw new Error(`No scopes for ${new URL(url).origin} on this connection; call request_scopes.`);
        const res = await conn.apiFetch(resource, method, url, { body, headers: { ...(content_type ? { 'Content-Type': content_type } : {}), ...headers } });
        const text = await res.text();
        return { status: res.status, contentType: res.headers.get('content-type'), body: text.length > 200000 ? `${text.slice(0, 200000)}\n[truncated]` : text };
      },
    );
}
