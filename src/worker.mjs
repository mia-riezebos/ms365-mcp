import { OAuthProvider, AuthorizationError, CimdFetchError, authorizationErrorRedirect } from '@cloudflare/workers-oauth-provider';
import { WebStandardStreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js';
import { createMcpServer } from './tools/index.mjs';
import { graphClient } from './graph.mjs';
import { BASE_SCOPES, union } from './scopes.mjs';
import {
  accessToken, accountOf, addScopes, authorizeUrl, createConnection, createScopeRequest, deleteScopeRequest,
  loadConnection, loadScopeRequest, pkcePair, redeemCode, saveScopeRequest,
} from './entra.mjs';

const DAY = 24 * 60 * 60;
const escape = (v) => String(v).replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
const STYLE = `:root{color-scheme:light dark;--bg:#f4f5f7;--card:#fff;--fg:#1b1d21;--muted:#5c6370;--line:#e3e5e9;--accent:#0f6cbd;--accent-fg:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#141518;--card:#1d1f23;--fg:#e8eaed;--muted:#9aa0a8;--line:#2e3137;--accent:#479ef5;--accent-fg:#0b0c0e}}
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;background:var(--bg);color:var(--fg);font:15px/1.55 system-ui,-apple-system,"Segoe UI",sans-serif;padding:1.5rem}
main{width:100%;max-width:30rem;background:var(--card);border:1px solid var(--line);border-radius:14px;padding:2rem;box-shadow:0 8px 30px rgba(0,0,0,.06)}
.brand{display:flex;align-items:center;gap:.6rem;color:var(--muted);font-size:.85rem;font-weight:600;letter-spacing:.02em;margin-bottom:1.25rem}
.mark{display:grid;place-items:center;width:28px;height:28px;border-radius:7px;background:var(--fg);color:var(--card);font:700 11px ui-monospace,Menlo,monospace}.brand b{color:var(--fg);font-weight:700}.brand small{display:block;font-weight:400;letter-spacing:0}
h1{font-size:1.35rem;line-height:1.3;margin:0 0 .75rem}p{margin:.6rem 0;color:var(--muted)}strong{color:var(--fg)}
code{font:.85em ui-monospace,SFMono-Regular,Menlo,monospace;background:var(--bg);border:1px solid var(--line);border-radius:5px;padding:.05em .35em}
ul{padding-left:1.2rem;color:var(--muted)}li{margin:.25rem 0}
.facts{border:1px solid var(--line);border-radius:10px;padding:.25rem 1rem;margin:1rem 0}.facts div{display:flex;justify-content:space-between;gap:1rem;padding:.55rem 0;border-bottom:1px solid var(--line);font-size:.9rem}.facts div:last-child{border:0}.facts span{color:var(--muted)}
.warn{background:rgba(255,185,0,.12);border:1px solid rgba(255,185,0,.45);border-radius:10px;padding:.6rem .85rem;color:var(--fg)}
.actions{display:flex;gap:.6rem;margin-top:1.5rem}button{flex:1;font:inherit;font-weight:600;padding:.7rem 1rem;border-radius:9px;border:1px solid var(--line);background:transparent;color:var(--fg);cursor:pointer}
button.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-fg)}
button.ms{display:flex;align-items:center;justify-content:center;gap:.6rem;background:#fff;color:#5e5e5e;border-color:#8c8c8c;border-radius:4px}button.ms svg{flex:none}button:hover{filter:brightness(1.08)}
.foot{margin-top:1.5rem;font-size:.8rem}a{color:var(--accent)}`;
const html = (body, init = {}) => {
  const headers = new Headers(init.headers);
  headers.set('Content-Type', 'text/html; charset=utf-8');
  return new Response(`<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>${STYLE}</style><main><div class="brand"><span class="mark">mcp</span><div><b>ms365-mcp</b><small>Open-source MCP server for Microsoft 365, not run by Microsoft</small></div></div>${body}</main></html>`, { ...init, headers });
};
const errorPage = (title, detail, status = 400) =>
  html(`<title>${escape(title)}</title><h1>${escape(title)}</h1><p>${escape(detail)}</p><p class="foot">Go back to your agent and connect again. If it keeps failing, share this message with whoever runs this server.</p>`, { status });
const text = (msg, status = 400) => new Response(msg, { status, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
const list = (v) => (v || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean);
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// Why an account may not use this server, or undefined. ALLOWED_TENANTS / ALLOWED_USERS narrow it; a tenant id
// in MS365_TENANT_ID allows only that tenant.
function refusal(env, { tenantId, upn }) {
  const tenants = list(env.ALLOWED_TENANTS);
  if (!tenants.length && GUID.test(env.MS365_TENANT_ID || '')) tenants.push(env.MS365_TENANT_ID.toLowerCase());
  if (tenants.length && !tenants.includes(String(tenantId).toLowerCase())) return 'This Microsoft account is not in a tenant allowed on this server.';
  const users = list(env.ALLOWED_USERS);
  if (users.length && !users.includes(upn)) return `${upn} is not allowed to use this server.`;
}
const LOOPBACK = ['localhost', '127.0.0.1', '[::1]'];
// Any HTTPS, loopback or app-scheme redirect unless ALLOWED_REDIRECT_HOSTS narrows it; the consent page names the host either way.
function redirectAllowed(env, redirectUri) {
  const url = new URL(redirectUri);
  const hosts = (env.ALLOWED_REDIRECT_HOSTS || '').split(',').map((s) => s.trim()).filter(Boolean);
  if (hosts.length) return hosts.includes(url.hostname);
  if (url.protocol === 'https:') return true;
  if (url.protocol === 'http:') return LOOPBACK.includes(url.hostname);
  return !['javascript:', 'data:', 'file:', 'blob:'].includes(url.protocol);
}

const mcpHandler = {
  async fetch(request, env, ctx) {
    if (request.method !== 'POST') return text('Method not allowed', 405);
    const { connId } = ctx.props;
    const record = await loadConnection(env, connId);
    if (!record) return text('This connection has no Microsoft account; reconnect the MCP server.', 401);
    const origin = new URL(request.url).origin;
    const conn = {
      granted: record.scopes,
      upn: record.upn,
      requestScopes: async (scopes, reason) => `${origin}/scopes/${await createScopeRequest(env, connId, scopes, reason)}`,
      apiFetch: async (resource, method, url, { body, headers }) =>
        fetch(url, { method, body, headers: { ...headers, Authorization: `Bearer ${await accessToken(env, connId, resource)}` } }),
    };
    const server = createMcpServer(graphClient(() => accessToken(env, connId)), conn);
    const transport = new WebStandardStreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
    await server.connect(transport);
    return transport.handleRequest(request);
  },
};

const MS_LOGO = '<svg width="20" height="20" viewBox="0 0 21 21" aria-hidden="true"><path fill="#f25022" d="M1 1h9v9H1z"/><path fill="#7fba00" d="M11 1h9v9h-9z"/><path fill="#00a4ef" d="M1 11h9v9H1z"/><path fill="#ffb900" d="M11 11h9v9h-9z"/></svg>';

function consentPage(details, handle, host) {
  const app = escape(details.clientName || 'An app');
  return html(`<title>Connect ${app} to ms365-mcp</title>
<h1>Let ${app} use your Microsoft 365 account through ${escape(host)}?</h1>
<p>This server (<strong>${escape(host)}</strong>) passes ${app}'s requests to Microsoft 365 as you. Signing in with Microsoft first only tells it who you are. When ${app} needs more, like reading mail or a calendar, Microsoft asks you again and lists exactly what it wants.</p>
<div class="facts"><div><span>App</span><strong>${app}</strong></div>
<div><span>Publisher</span><strong>${details.clientDomain ? escape(details.clientDomain) : 'Not verified'}</strong></div>
<div><span>Access goes to</span><strong>${escape(details.redirectHost)}</strong></div></div>
${details.redirectIsLoopback ? '<p class="warn">This sends access to an app on your computer. Continue only if you just started connecting from it.</p>' : ''}
<form method="post"><input type="hidden" name="handle" value="${escape(handle)}">
<div class="actions"><button name="decision" value="deny">Cancel</button><button class="ms" name="decision" value="approve">${MS_LOGO}Sign in with Microsoft</button></div></form>
<p class="foot">Source: <a href="https://github.com/mia-riezebos/ms365-mcp">github.com/mia-riezebos/ms365-mcp</a></p>`);
}

async function authorize(request, env) {
  const oauth = env.OAUTH_PROVIDER;
  if (request.method === 'GET') {
    const authRequest = await oauth.parseAuthRequest(request);
    if (!redirectAllowed(env, authRequest.redirectUri)) return text(`Redirect to ${new URL(authRequest.redirectUri).origin} is not allowed.`);
    const details = await oauth.describeConsent(authRequest);
    const consent = await oauth.beginConsent(authRequest);
    return withHeaders(consentPage(details, consent.handle, new URL(request.url).host), consent.headers);
  }
  if (request.method !== 'POST') return text('Method not allowed', 405);
  const form = await request.formData();
  const handle = String(form.get('handle'));
  if (form.get('decision') !== 'approve') {
    const denied = await oauth.denyConsent(request, handle);
    return new Response(null, { status: 302, headers: denied.headers });
  }
  const approved = await oauth.approveConsent(request, handle);
  const { verifier, challenge } = await pkcePair();
  const { state, headers } = await oauth.beginUpstream(approved.request, { data: { verifier }, headers: approved.headers });
  headers.set('Location', authorizeUrl(env, request, { state, codeChallenge: challenge, scopes: BASE_SCOPES }));
  return new Response(null, { status: 302, headers });
}

function withHeaders(response, headers) {
  for (const [k, v] of headers) response.headers.set(k, v);
  return response;
}

async function callback(request, env) {
  const params = new URL(request.url).searchParams;
  if (params.get('state')?.startsWith(SCOPE_STATE)) return scopeCallback(request, env, params);
  const oauth = env.OAUTH_PROVIDER;
  const { request: original, data, headers } = await oauth.finishUpstream(request);
  if (params.get('error')) {
    console.warn('sign-in returned error', params.get('error'), params.get('error_description'));
    const description = params.get('error_description')?.split('\r\n')[0];
    if (params.get('error') !== 'access_denied') return withHeaders(errorPage('Microsoft sign-in failed', description || params.get('error'), 502), headers);
    headers.set('Location', authorizationErrorRedirect(original, 'access_denied', description));
    return new Response(null, { status: 302, headers });
  }
  const tokens = await redeemCode(env, request, params.get('code'), data.verifier, BASE_SCOPES);
  const account = accountOf(tokens);
  const refused = refusal(env, account);
  if (refused) return errorPage('This account cannot connect', refused, 403);
  console.log('signed in', account.tenantId);
  const connId = await createConnection(env, account, tokens);
  const { redirectTo } = await oauth.completeAuthorization({
    request: original,
    userId: account.userId,
    metadata: { upn: account.upn },
    scope: original.scope,
    props: { connId, upn: account.upn },
  });
  headers.set('Location', redirectTo);
  return new Response(null, { status: 302, headers });
}

const SCOPE_STATE = 'scopes.';
const scopeList = (scopes) => `<ul>${scopes.map((s) => `<li><code>${escape(s)}</code></li>`).join('')}</ul>`;

// GET /scopes/<id>: the link request_scopes hands out. Sends the user to Microsoft for the connection's scopes plus the new ones.
async function scopeStart(request, env, id) {
  const req = await loadScopeRequest(env, id);
  const conn = req && (await loadConnection(env, req.connId));
  if (!conn) return html('<title>Link expired</title><h1>This link has expired</h1><p>Ask your agent to request the access again.</p>', { status: 404 });
  const { verifier, challenge } = await pkcePair();
  await saveScopeRequest(env, id, { ...req, verifier });
  return Response.redirect(authorizeUrl(env, request, { state: `${SCOPE_STATE}${id}`, codeChallenge: challenge, scopes: union(conn.scopes, req.scopes), loginHint: conn.upn }), 302);
}

async function scopeCallback(request, env, params) {
  const id = params.get('state').slice(SCOPE_STATE.length);
  const req = await loadScopeRequest(env, id);
  const conn = req?.verifier && (await loadConnection(env, req.connId));
  if (!conn) return html('<title>Link expired</title><h1>This link has expired</h1><p>Ask your agent to request the access again.</p>', { status: 404 });
  await deleteScopeRequest(env, id);
  if (params.get('error'))
    return html(`<title>Not granted</title><h1>Nothing was granted</h1><p>${escape(params.get('error_description')?.split('\r\n')[0] || params.get('error'))}</p>`);
  const asked = union(conn.scopes, req.scopes);
  const tokens = await redeemCode(env, request, params.get('code'), req.verifier, asked);
  const account = accountOf(tokens);
  if (account.userId !== conn.userId || account.tenantId !== conn.tenantId)
    return errorPage('Wrong Microsoft account', `This connection belongs to ${conn.upn}; sign in with that account.`, 403);
  const { refused } = await addScopes(env, req.connId, conn, tokens, req.scopes);
  const granted = req.scopes.filter((s) => !refused.includes(s));
  return html(`<title>Access granted</title><h1>${granted.length ? 'Access granted' : 'Nothing was granted'}</h1>
${granted.length ? `<p>Granted:</p>${scopeList(granted)}` : ''}${refused.length ? `<p>Not granted:</p>${scopeList(refused)}` : ''}
<p>You can close this tab and go back to your agent.</p>`);
}

const defaultHandler = {
  async fetch(request, env) {
    const { pathname, origin } = new URL(request.url);
    try {
      if (pathname === '/authorize') return await authorize(request, env);
      if (pathname === '/auth/callback') return await callback(request, env);
      if (pathname.startsWith('/scopes/')) return await scopeStart(request, env, pathname.slice('/scopes/'.length));
    } catch (error) {
      console.error(pathname, error?.stack || error);
      if (error instanceof AuthorizationError && error.redirectTo) return Response.redirect(error.redirectTo, 302);
      if (error instanceof AuthorizationError) return errorPage('Could not connect', error.description);
      if (error instanceof CimdFetchError) return errorPage('Could not connect', 'This app could not be verified.');
      return errorPage('Sign-in failed', error?.message || String(error), 500);
    }
    if (pathname === '/logout') return text('Signed out', 200);
    if (pathname === '/healthz') return text('ok', 200);
    if (pathname === '/')
      return html(`<title>Microsoft 365 MCP</title><h1>Microsoft 365 MCP</h1>
<p>Remote MCP server for Microsoft 365 as the signed-in user: mail (incl. shared mailboxes), calendar, contacts, Teams, Planner, To Do, OneDrive/SharePoint and OneNote, plus raw Graph calls. Agents start with no access and ask for the Microsoft permissions they need with <code>request_scopes</code>; each request is a Microsoft sign-in you approve.</p>
<p>Source: <a href="https://github.com/mia-riezebos/ms365-mcp">github.com/mia-riezebos/ms365-mcp</a></p>
<p>MCP endpoint: <code>${escape(origin)}/mcp</code>. Connect it as an OAuth MCP server; connecting signs you in with Microsoft.</p>`);
    return text('Not found', 404);
  },
};

const providerOptions = (env) => ({
  apiRoute: '/mcp',
  apiHandler: mcpHandler,
  defaultHandler,
  authorizeEndpoint: '/authorize',
  tokenEndpoint: '/oauth/token',
  clientRegistrationEndpoint: '/oauth/register',
  scopesSupported: ['ms365'],
  accessTokenTTL: 60 * 60,
  // Grants live while the client keeps refreshing, matching Entra's 90-day sliding refresh token window.
  refreshTokenTTL: 90 * DAY,
  refreshTokenIdleTTL: 90 * DAY,
  clientRegistrationTTL: undefined,
  resourceMetadata: { resource: `${env.PUBLIC_BASE_URL}/mcp`, resource_name: 'Microsoft 365' },
});

// The resource URL comes from env, so the provider is built on the first request.
let provider;
export default {
  fetch(request, env, ctx) {
    provider ??= new OAuthProvider(providerOptions(env));
    return provider.fetch(request, env, ctx);
  },
};
